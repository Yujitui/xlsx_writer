//! Excel 工作簿导出管理器模块
//!
//! 本模块定义了 Excel 导出的核心容器结构 Workbook，负责：
//! - 管理全局样式配置（样式池）
//! - 维护工作表导出任务队列
//! - 协调多工作表的统一导出流程
use crate::cell::Cell;
use crate::error::XlsxError;
use crate::merge_factory::MergeFactory;
use crate::prelude::ReadSheet;
use crate::style_factory::StyleFactory;
use crate::style_library::StyleLibrary;
use crate::worksheet::WorkSheet;
use calamine::{open_workbook, DataType, Reader, Xlsx};
use polars::prelude::*;
use rust_xlsxwriter::*;
use std::collections::HashMap;
use std::error::Error;
use std::fmt;
use std::io::{BufReader, Read};
use std::path::Path;

/// Excel 工作簿导出管理器
///
/// 作为 Excel 导出系统的顶层容器，负责协调多个工作表的统一导出。
/// 通过样式池和任务队列的管理，实现高效、一致的 Excel 文件生成。
pub struct Workbook {
    /// 全局样式配置池（私有字段）
    ///
    /// 存储预定义的样式格式，通过字符串标签进行引用。
    /// 使用 HashMap 实现 O(1) 的样式查找性能。
    styles: HashMap<String, Format>,
    /// 工作表导出任务队列（私有字段）
    ///
    /// 按照 Vec 中的顺序进行工作表导出，索引 0 对应 Excel 中的第一个工作表。
    /// 每个工作表包含完整的数据和样式信息。
    sheets: Vec<WorkSheet>,
}

// ============================================================================
// 1. 核心访问 API
// ============================================================================

impl Workbook {
    /// 获取工作簿中工作表的数量
    pub fn len(&self) -> usize {
        self.sheets.len()
    }

    /// 检查工作簿是否为空
    pub fn is_empty(&self) -> bool {
        self.sheets.is_empty()
    }

    /// 按索引获取工作表名称
    ///
    /// # 参数
    /// * `index` - 工作表索引
    ///
    /// # 返回值
    /// * `Some(&str)` - 工作表名称
    /// * `None` - 索引无效
    pub fn sheet_name(&self, index: usize) -> Option<&str> {
        self.sheets.get(index).map(|sheet| sheet.name.as_str())
    }

    /// 通过工作表名称查找工作表
    ///
    /// 根据指定的名称在工作簿中查找对应的工作表。
    ///
    /// # 参数
    /// * `name` - 要查找的工作表名称
    ///
    /// # 返回值
    /// * `Some(&WorkSheet)` - 找到匹配的工作表，返回其引用
    /// * `None` - 未找到指定名称的工作表
    pub fn sheet(&self, name: &str) -> Option<&WorkSheet> {
        self.sheets.iter().find(|sheet| sheet.name == name)
    }

    /// 通过工作表名称查找可变引用
    ///
    /// # 参数
    /// * `name` - 要查找的工作表名称
    ///
    /// # 返回值
    /// * `Some(&mut WorkSheet)` - 找到匹配的工作表，返回其可变引用
    /// * `None` - 未找到指定名称的工作表
    pub fn sheet_mut(&mut self, name: &str) -> Option<&mut WorkSheet> {
        self.sheets.iter_mut().find(|sheet| sheet.name == name)
    }

    /// 按索引获取工作表的样式映射（简化版）
    ///
    /// 返回样式映射的直接引用，如果工作表不存在则返回 None。
    ///
    /// # 参数
    /// * `index` - 工作表索引
    ///
    /// # 返回值
    /// * `Some(&HashMap<(u32, u16), Arc<str>>)` - 存在且有样式映射
    /// * `None` - 工作表不存在或无样式映射
    pub fn style_map(&self, index: usize) -> Option<&HashMap<(u32, u16), Arc<str>>> {
        self.sheets
            .get(index)
            .and_then(|sheet| sheet.style_map.as_ref())
    }
}

// ============================================================================
// 2. 构造与配置
// ============================================================================

impl Workbook {
    /// 创建并初始化新的 Workbook 实例
    ///
    /// 工厂方法模式实现，负责 Workbook 对象的构造和基础环境初始化。
    /// 通过跨平台字体适配和预置样式，为 Excel 导出提供一致的默认体验。
    pub fn new() -> Result<Self, Box<dyn Error>> {
        let mut styles = HashMap::new();

        // 1. 跨平台字體自適應適配：
        // 利用編譯時宏 cfg! 確定目標系統。Excel 文件不嵌入字體，僅存儲名稱，
        // 故需確保指定的名稱在對應系統中是可被識別的標準名稱。
        let font_name = if cfg!(target_os = "windows") {
            "Microsoft YaHei" // Windows: 微软雅黑
        } else if cfg!(target_os = "macos") {
            "PingFang SC" // macOS: 苹方
        } else {
            "sans-serif" // Linux 或其他: 通用無襯線字體
        };

        // 2. 定義「標準樣式」：
        // 這是報表中絕大多數單元格的底層模板，包含了邊框、字體和對齊方式。
        let standard_fmt = Format::new()
            .set_font_name(font_name) // 設置微軟雅黑
            .set_font_size(11) // 設置常用字號
            .set_border(FormatBorder::Thin) // 設置細邊框（四周）
            .set_align(FormatAlign::Center) // 水平居中
            .set_align(FormatAlign::VerticalCenter); // 垂直居中

        // 3. 快速派生「表頭樣式」：
        // 通過 clone() 繼承 standard_fmt 的屬性，僅對差異項（加粗、背景）進行覆蓋。
        // 這種鏈式調用確保了表頭與數據行在字體、邊框寬度上完全對齊。
        let header_fmt = standard_fmt
            .clone()
            .set_bold()
            .set_background_color(Color::RGB(0xBFBFBF));

        // 4. 將初始化樣式存入緩存池：
        // "default" 用於普通單元格保底，"header" 用於標題行。
        styles.insert("default".to_string(), standard_fmt);
        styles.insert("header".to_string(), header_fmt);

        Ok(Self {
            styles,
            sheets: vec![],
        })
    }

    // ============================================================================
    // 3. I/O 操作（读取和保存）
    // ============================================================================

    /// 从Excel文件读取多个工作表数据
    ///
    /// 支持自动格式检测和回退机制。首先根据文件扩展名选择读取方法，
    /// 如果首选方法失败，会自动尝试备选格式。
    ///
    /// # 参数
    /// * `path` - Excel文件的完整路径
    /// * `read_sheets` - 工作表读取配置列表，如果为空则读取所有工作表
    ///
    /// # 返回值
    /// 返回当前Workbook实例，便于链式调用
    pub fn read(mut self, path: &str, read_sheets: Vec<ReadSheet>) -> Self {
        // 按后缀确定首选方法
        let preferred_is_xls = path.to_lowercase().ends_with(".xls");

        // 尝试首选方法
        let primary_result = if preferred_is_xls {
            self.try_read_as_xls(path, &read_sheets)
        } else {
            self.try_read_as_xlsx(path, &read_sheets)
        };

        match primary_result {
            Ok(_) => self,
            Err(primary_err) => {
                eprintln!(
                    "首选读取方法失败（后缀 {}）: {}",
                    if preferred_is_xls { ".xls" } else { ".xlsx" },
                    primary_err
                );

                // 清空已添加的工作表（为未来扩展预留）
                self.sheets.clear();

                // 尝试备选方法
                let fallback_result = if preferred_is_xls {
                    eprintln!("尝试回退到 .xlsx 格式读取...");
                    self.try_read_as_xlsx(path, &read_sheets)
                } else {
                    eprintln!("尝试回退到 .xls 格式读取...");
                    self.try_read_as_xls(path, &read_sheets)
                };

                match fallback_result {
                    Ok(_) => {
                        eprintln!("警告: 文件格式与后缀 '{}' 不符", path);
                        self
                    }
                    Err(fallback_err) => {
                        eprintln!("备选读取方法也失败: {}", fallback_err);
                        eprintln!("无法读取文件: {}", path);
                        self
                    }
                }
            }
        }
    }

    // ------------------------------------------------------------------------
    // 私有辅助方法
    // ------------------------------------------------------------------------

    /// 去除重复的 ReadSheet，保持原始顺序
    fn dedup_read_sheets(read_sheets: Vec<ReadSheet>) -> Vec<ReadSheet> {
        let mut seen = std::collections::HashSet::new();
        read_sheets
            .into_iter()
            .filter(|sheet| seen.insert(sheet.clone()))
            .collect()
    }

    /// 尝试读取为 .xlsx 格式
    fn try_read_as_xlsx(&mut self, path: &str, read_sheets: &Vec<ReadSheet>) -> Result<(), String> {
        use calamine::Data;
        use chrono::NaiveDate;

        let mut excel: Xlsx<BufReader<std::fs::File>> = match open_workbook(Path::new(path)) {
            Ok(excel) => excel,
            Err(e) => return Err(format!("无法打开文件: {}", e)),
        };

        let read_sheets = Self::dedup_read_sheets(read_sheets.clone());

        let read_sheets = if read_sheets.is_empty() {
            let sheet_names = excel.sheet_names();
            if sheet_names.is_empty() {
                return Err("Excel文件中没有工作表".to_string());
            }
            sheet_names
                .iter()
                .map(|name| ReadSheet::new(name.to_string()))
                .collect()
        } else {
            read_sheets
        };

        if read_sheets.is_empty() {
            return Ok(());
        }

        for sheet_config in &read_sheets {
            // 直接使用 calamine 读取数据并转换为 cells
            let result = (|| -> Result<WorkSheet, XlsxError> {
                let sheet_names = excel.sheet_names();

                // 确定目标工作表
                let target_sheet = if sheet_names.contains(&sheet_config.sheet_name) {
                    &sheet_config.sheet_name
                } else {
                    return Err(XlsxError::SheetNotFound(sheet_config.sheet_name.clone()));
                };

                let skip_rows = sheet_config.skip_rows.unwrap_or(0);
                let force_string_cols = sheet_config.force_string_cols.clone().unwrap_or_default();

                // 获取工作表数据
                let range = excel
                    .worksheet_range(target_sheet)
                    .map_err(|e| XlsxError::CalamineError(e))?;

                let mut rows = range.rows().skip(skip_rows);

                // 提取表头
                let headers: Vec<String> = {
                    let first_row = rows.next().ok_or(XlsxError::MissingHeaderRow)?;
                    first_row
                        .iter()
                        .enumerate()
                        .map(|(idx, cell)| {
                            let raw = cell.to_string();
                            if raw.trim().is_empty() {
                                format!("column_{}", idx)
                            } else {
                                raw
                            }
                        })
                        .collect()
                };

                // 构建 cells
                let mut cells: Vec<Vec<Option<Cell>>> = Vec::new();

                // 第 0 行：表头
                cells.push(
                    headers
                        .iter()
                        .map(|h| Some(Cell::Text(h.clone())))
                        .collect(),
                );

                // 数据行
                for row in rows {
                    let mut cell_row: Vec<Option<Cell>> = Vec::new();
                    for (col_idx, cell) in row.iter().enumerate() {
                        let cell_value = if force_string_cols.contains(&headers[col_idx]) {
                            Some(Cell::Text(cell.to_string()))
                        } else {
                            match cell {
                                Data::Int(val) => Some(Cell::Number(*val as f64)),
                                Data::Float(val) => Some(Cell::Number(*val)),
                                Data::String(val) => Some(Cell::Text(val.clone())),
                                Data::Bool(val) => Some(Cell::Boolean(*val)),
                                Data::Empty => None,
                                _ => {
                                    if let Some(date) = cell.as_date() {
                                        // Excel 日期转换为序列号
                                        let excel_epoch =
                                            NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
                                        let days = (date - excel_epoch).num_days();
                                        Some(Cell::Number(days as f64))
                                    } else {
                                        Some(Cell::Text(cell.to_string()))
                                    }
                                }
                            }
                        };
                        cell_row.push(cell_value);
                    }
                    cells.push(cell_row);
                }

                WorkSheet::new(cells, sheet_config.sheet_name.clone(), None, None)
            })();

            match result {
                Ok(worksheet) => {
                    if self.sheet(worksheet.name.as_str()).is_none() {
                        self.sheets.push(worksheet);
                    } else {
                        eprintln!("sheet '{}' 与现有表格重复", worksheet.name);
                    }
                }
                Err(e) => {
                    eprintln!("读取sheet '{}' 失败: {}", sheet_config.sheet_name, e);
                }
            }
        }

        Ok(())
    }

    /// 将特殊区域写入 worksheet
    ///
    /// # 参数
    /// * `worksheet` - 目标工作表
    /// * `region` - 要写入的特殊区域
    /// * `start_row` - 起始行号
    /// * `fallback_fmt` - 默认样式
    ///
    /// # 返回值
    /// 返回写入后的下一行号
    fn write_region_to_worksheet(
        &self,
        worksheet: &mut Worksheet,
        region: &crate::sheet_region::SheetRegion,
        start_row: u32,
        fallback_fmt: &Format,
    ) -> Result<u32, Box<dyn Error>> {
        let default_fmt = self.styles.get("default");
        let mut current_row = start_row;

        // 写入区域数据
        for (rel_row, row_data) in region.data.iter().enumerate() {
            for (rel_col, cell_opt) in row_data.iter().enumerate() {
                let r = start_row + rel_row as u32;
                let c = rel_col as u16;

                // 获取单元格样式
                let cell_fmt = region
                    .style_map
                    .as_ref()
                    .and_then(|m| m.get(&(rel_row as u32, c)))
                    .and_then(|name| self.styles.get(name.as_ref()))
                    .or(default_fmt)
                    .unwrap_or(fallback_fmt);

                // 写入单元格
                if let Some(cell) = cell_opt {
                    match cell {
                        crate::Cell::Text(s) => {
                            worksheet.write_with_format(r, c, s.as_str(), cell_fmt)?;
                        }
                        crate::Cell::Number(n) => {
                            worksheet.write_number_with_format(r, c, *n, cell_fmt)?;
                        }
                        crate::Cell::Boolean(b) => {
                            worksheet.write_boolean_with_format(r, c, *b, cell_fmt)?;
                        }
                    }
                } else {
                    // 空单元格，写入 blank 保持样式
                    worksheet.write_blank(r, c, cell_fmt)?;
                }
            }
            current_row = start_row + rel_row as u32 + 1;
        }

        // 处理合并区域
        if let Some(ref merge_ranges) = region.merge_ranges {
            for merge_range in merge_ranges {
                worksheet.merge_range(
                    start_row + merge_range.0, // first_row
                    merge_range.1,             // first_col
                    start_row + merge_range.2, // last_row
                    merge_range.3,             // last_col
                    "",                        // 数据为空白字符串
                    fallback_fmt,              // 使用默认样式
                )?;
            }
        }

        Ok(current_row)
    }

    // ------------------------------------------------------------------------
    // XLS 格式专用方法（BIFF8）
    // ------------------------------------------------------------------------

    /// 尝试读取为 .xls 格式
    fn try_read_as_xls(&mut self, path: &str, read_sheets: &Vec<ReadSheet>) -> Result<(), String> {
        use cfb::CompoundFile;
        use std::fs::File;
        use std::io::{BufReader, Cursor, Read};

        // 打开文件并读取到缓冲区
        let file = File::open(path).map_err(|e| format!("无法打开文件: {}", e))?;
        let mut buf_reader = BufReader::new(file);
        let mut buffer = Vec::new();
        buf_reader
            .read_to_end(&mut buffer)
            .map_err(|e| format!("无法读取文件: {}", e))?;

        // 解析 CFB 结构
        let mut cursor = Cursor::new(buffer);
        let mut cfb = CompoundFile::open(&mut cursor).map_err(|e| format!("CFB 错误: {}", e))?;

        // 打开 Workbook 流
        let workbook_stream = cfb
            .open_stream("Workbook")
            .map_err(|_| "Workbook stream not found".to_string())?;

        // 读取所有字节
        let bytes: Vec<u8> = workbook_stream
            .bytes()
            .collect::<Result<Vec<_>, _>>()
            .map_err(|e| format!("读取流错误: {}", e))?;

        // 解析工作簿
        let mut reader = Cursor::new(bytes);
        let (sheet_names, sheets_data, sst) = Self::parse_xls_workbook(&mut reader)?;

        // 去重 read_sheets
        let read_sheets = Self::dedup_read_sheets(read_sheets.clone());

        // 确定要处理的工作表
        let sheets_to_process: Vec<(usize, Option<&ReadSheet>)> = if read_sheets.is_empty() {
            // 读取所有表
            (0..sheets_data.len()).map(|idx| (idx, None)).collect()
        } else {
            // 按 read_sheets 筛选
            read_sheets
                .iter()
                .filter_map(|rs| {
                    sheet_names
                        .iter()
                        .position(|name| name == &rs.sheet_name)
                        .map(|idx| (idx, Some(rs)))
                })
                .collect()
        };

        if sheets_to_process.is_empty() {
            return Err("没有匹配的工作表".to_string());
        }

        // 转换每个工作表
        for (idx, read_sheet) in sheets_to_process {
            if idx >= sheets_data.len() {
                continue;
            }
            let sheet_name = &sheet_names[idx];
            let rows = &sheets_data[idx];

            match Self::convert_to_worksheet(sheet_name, rows, read_sheet, &sst) {
                Ok(worksheet) => {
                    if self.sheet(worksheet.name.as_str()).is_none() {
                        self.sheets.push(worksheet);
                    } else {
                        eprintln!("sheet '{}' 与现有表格重复", worksheet.name);
                    }
                }
                Err(e) => {
                    eprintln!("转换工作表 '{}' 失败: {}", sheet_name, e);
                }
            }
        }

        Ok(())
    }

    /// 解析 XLS 工作簿，返回工作表名称、行数据和 SST
    fn parse_xls_workbook<R: Read>(
        reader: &mut R,
    ) -> Result<
        (
            Vec<String>,
            Vec<Vec<Vec<Option<crate::Cell>>>>,
            crate::xls_records::SharedStringTable,
        ),
        String,
    > {
        use crate::xls_records::{
            BlankRecord, BoFRecord, BoolErrRecord, BoundSheetRecord, ContinueRecord,
            DimensionsRecord, EofRecord, FormulaRecord, LabelSSTRecord, MulBlankRecord,
            MulRkRecord, NumberRecord, ParsableRecord, ParseState, RKRecord, RowRecord,
            SSTRecordData,
        };
        use byteorder::{LittleEndian, ReadBytesExt};

        let mut state = ParseState::new();

        loop {
            // 读取记录头部
            let record_type = match reader.read_u16::<LittleEndian>() {
                Ok(rt) => rt,
                Err(_) => break,
            };

            let length = match reader.read_u16::<LittleEndian>() {
                Ok(len) => len as usize,
                Err(_) => break,
            };

            // 读取记录数据
            let mut data = vec![0u8; length];
            if let Err(e) = reader.read_exact(&mut data) {
                eprintln!("Warning: Failed to read record data: {}", e);
                break;
            }

            // 匹配记录类型并应用
            let result: Result<(), String> = (|| match record_type {
                BoFRecord::RECORD_ID => BoFRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                EofRecord::RECORD_ID => {
                    EofRecord::parse(&data)
                        .map_err(|e| e.to_string())?
                        .apply(&mut state)
                        .map_err(|e| e.to_string())?;
                    if state.is_complete {
                        return Ok(());
                    }
                    Ok(())
                }
                SSTRecordData::RECORD_ID => SSTRecordData::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                ContinueRecord::RECORD_ID => ContinueRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                BoundSheetRecord::RECORD_ID => BoundSheetRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                RowRecord::RECORD_ID => RowRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                DimensionsRecord::RECORD_ID => DimensionsRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                NumberRecord::RECORD_ID => NumberRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                RKRecord::RECORD_ID => RKRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                BlankRecord::RECORD_ID => BlankRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                LabelSSTRecord::RECORD_ID => LabelSSTRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                MulRkRecord::RECORD_ID => MulRkRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                MulBlankRecord::RECORD_ID => MulBlankRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                BoolErrRecord::RECORD_ID => BoolErrRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                FormulaRecord::RECORD_ID => FormulaRecord::parse(&data)
                    .map_err(|e| e.to_string())?
                    .apply(&mut state)
                    .map_err(|e| e.to_string()),
                _ => Ok(()),
            })();

            if let Err(e) = result {
                eprintln!(
                    "Warning: Error processing record 0x{:04X}: {}",
                    record_type, e
                );
                // Check if we need to break early due to completion
                if record_type == EofRecord::RECORD_ID && state.is_complete {
                    break;
                }
            }
        }

        // Finalize SST
        if let Some(parser) = state.sst_parser.take() {
            if let Err(e) = parser.finish(&mut state.sst) {
                eprintln!("Warning: Failed to finish SST: {}", e);
            }
        }

        // Add current sheet if any
        if let Some(sheet) = state.current_sheet.take() {
            state.sheets.push(sheet);
        }

        // 提取工作表名称和行数据
        let sheet_names = state.sheet_names;
        let sheets_data: Vec<Vec<Vec<Option<crate::Cell>>>> =
            state.sheets.into_iter().map(|sheet| sheet.rows).collect();

        Ok((sheet_names, sheets_data, state.sst))
    }

    /// 将 XLS 行数据转换为 WorkSheet
    fn convert_to_worksheet(
        sheet_name: &str,
        rows: &[Vec<Option<crate::Cell>>],
        read_sheet: Option<&ReadSheet>,
        _sst: &crate::xls_records::SharedStringTable,
    ) -> Result<WorkSheet, crate::error::XlsxError> {
        use crate::cell::Cell;

        let skip_rows = read_sheet.and_then(|r| r.skip_rows).unwrap_or(0);

        // 检查是否有足够的数据
        if rows.len() <= skip_rows + 1 {
            return Err(crate::error::XlsxError::EmptyDataFrame);
        }

        let header_row_idx = skip_rows;
        let max_cols = rows.iter().map(|row| row.len()).max().unwrap_or(0);

        if max_cols == 0 {
            return Err(crate::error::XlsxError::EmptyDataFrame);
        }

        // 构建 cells
        let mut cells: Vec<Vec<Option<Cell>>> = Vec::new();

        // 表头行
        let header_row = &rows[header_row_idx];
        let mut used_names: std::collections::HashMap<String, usize> =
            std::collections::HashMap::new();

        let header_cells: Vec<Option<Cell>> = (0..max_cols)
            .map(|col_idx| {
                let base_name = if let Some(Some(cell)) = header_row.get(col_idx) {
                    match cell {
                        crate::Cell::Text(s) if !s.is_empty() => s.clone(),
                        crate::Cell::Number(n) => n.to_string(),
                        crate::Cell::Boolean(b) => b.to_string(),
                        _ => format!("Column_{}", col_idx),
                    }
                } else {
                    format!("Column_{}", col_idx)
                };

                let count = used_names.entry(base_name.clone()).or_insert(0);
                *count += 1;
                let final_name = if *count > 1 {
                    format!("{}_{}", base_name, *count)
                } else {
                    base_name
                };

                Some(Cell::Text(final_name))
            })
            .collect();
        cells.push(header_cells);

        // 数据行
        for row_idx in (skip_rows + 1)..rows.len() {
            let row = &rows[row_idx];
            let cell_row: Vec<Option<Cell>> = (0..max_cols)
                .map(|col_idx| {
                    row.get(col_idx)
                        .and_then(|opt| opt.as_ref())
                        .map(|cell| match cell {
                            crate::Cell::Text(s) => Cell::Text(s.clone()),
                            crate::Cell::Number(n) => Cell::Number(*n),
                            crate::Cell::Boolean(b) => Cell::Boolean(*b),
                        })
                })
                .collect();
            cells.push(cell_row);
        }

        WorkSheet::new(cells, sheet_name.to_string(), None, None)
    }

    /// 注册或更新工作簿中的样式定义
    ///
    /// 构建器模式方法，用于向工作簿的样式池中添加新的样式定义
    /// 或更新已存在的样式配置。通过链式调用支持流畅的样式配置体验。
    ///
    /// # 参数说明
    ///
    /// * `name`: 样式名称标识符
    ///   - 用于在样式映射中唯一标识样式
    ///   - 在单元格样式应用时作为引用键使用
    /// * `format`: 实际的样式格式对象
    ///   - `rust_xlsxwriter::Format` 类型的完整样式定义
    ///   - 包含字体、颜色、对齐、边框等所有样式属性
    pub fn set_style(mut self, name: &str, format: Format) -> Self {
        // 將字符串切片轉換為 String 並存入 HashMap。
        // HashMap 的 insert 特性確保了同名鍵值的自動更新。
        self.styles.insert(name.to_string(), format);

        // 返回修改後的對象，支持鏈式調用：exporter.set_style(...).set_style(...)
        self
    }

    /// 使用样式库扩展工作簿的样式配置
    ///
    /// 构建器模式方法，用于将外部定义的样式库集成到当前工作簿中。
    /// 通过批量注册样式库中的所有预定义样式，实现样式的统一管理和复用。
    ///
    /// # 参数说明
    ///
    /// * `library`: 对样式库的不可变引用
    ///   - 包含预定义的样式配置集合
    ///   - 通过 `build_formats()` 方法转换为实际的 Format 对象
    pub fn with_library(mut self, library: &StyleLibrary) -> Self {
        // 調用 Library 內部的批量轉換邏輯
        // 将样式库定义转换为实际的 Format 对象集合
        let formats = library.build_formats();

        for (name, fmt) in formats {
            // 利用現有的 set_style 實現，將 Format 存入 Workbook.styles
            // 通过已有的样式注册机制确保一致性和冲突处理
            self = self.set_style(&name, fmt);
        }

        self
    }

    /// 从 JSON 配置应用样式到工作簿
    ///
    /// 便捷方法，用于直接从 JSON 配置数据创建样式库并应用到当前工作簿。
    /// 通过两阶段处理（JSON 解析 → 样式库构建 → 样式应用）实现配置驱动的样式管理。
    ///
    /// # 参数说明
    ///
    /// * `value`: 对 JSON 配置的不可变引用
    ///   - 包含样式定义的结构化数据
    ///   - 格式需符合 StyleLibrary 的 JSON Schema
    ///
    /// # JSON 配置格式
    ///
    /// ## 基本结构
    /// ```json
    /// {
    ///   "styles": {
    ///     "header_bg": { "bg_color": "#BFBFBF", "bold": true, "border": "thin", "align": "center" },
    ///     "money_red": { "font_color": "#FF0000", "num_format": "#,##0.00" },
    ///     "default": { "font_name": "Microsoft YaHei", "font_size": 11 }
    ///   },
    /// }
    /// ```
    pub fn with_library_from_json(self, value: &serde_json::Value) -> Result<Self, Box<dyn Error>> {
        // 1. 在內部利用片段構建樣式庫
        // 通过 JSON 配置创建样式库实例
        // 此步骤会进行配置解析和验证
        let library = StyleLibrary::from_json(value)?;

        // 2. 調用前者進行物理注入
        // 利用已有的 with_library 方法应用样式库
        // 保持样式应用逻辑的一致性
        Ok(self.with_library(&library))
    }

    /// 向工作簿中添加新的工作表任务
    ///
    /// 构建器模式方法，负责将 DataFrame 数据封装为 WorkSheet 任务并添加到工作簿中。
    /// 通过智能的错误处理和自动修复机制，提供鲁棒的工作表添加体验。
    ///
    /// # 参数说明
    ///
    /// * `df`: 待导出的 DataFrame 数据源
    ///   - 包含实际的数据内容
    ///   - 支持所有权转移以提高性能
    /// * `name`: 可选的工作表名称
    ///   - None：使用默认名称 "Sheet N"
    ///   - Some(name)：使用提供的名称（会进行验证）
    /// * `style_map`: 可选的单元格样式映射
    ///   - None：不应用特殊样式
    ///   - Some(map)：应用指定的样式映射
    /// * `merge_ranges`: 可选的工作表合并区域范围列表
    ///   - None：不进行任何合并操作
    ///   - Some(ranges)：指定需要合并的单元格范围（每组四个参数表示起始行、起始列、结束行、结束列）
    ///
    /// # 返回值
    ///
    /// * `Ok(Workbook)` - 成功添加工作表后的工作簿实例
    /// * `Err(Box<dyn Error>)` - 添加过程中发生的错误
    pub fn insert(
        mut self,
        df: DataFrame,
        name: Option<String>,
        style_map: Option<HashMap<(u32, u16), Arc<str>>>,
        merge_ranges: Option<Vec<(u32, u16, u32, u16)>>,
    ) -> Result<Self, Box<dyn Error>> {
        // 1. 定義輔助閉包：封裝默認命名邏輯，確保命名的一致性與唯一性起點
        let get_default_name = |sheets_len: usize| format!("Sheet {}", sheets_len + 1);

        // 2. 初步確定名稱：優先使用用戶提供，否則生成默認名
        let final_name = name.unwrap_or_else(|| get_default_name(self.sheets.len()));

        // 3. 嘗試構建 WorkSheet 任務：
        // 這裡使用 match 進行細分錯誤處理。注意 df.clone() 是淺拷貝，開銷極低。
        // 這裡 clone() 是必須的，因為如果 InvalidName 發生，我們需要原始數據進行第二次嘗試。
        let task = match WorkSheet::from_dataframe(
            df.clone(),
            final_name.clone(),
            style_map.clone(),
            merge_ranges.clone(),
        ) {
            Ok(t) => t,
            // 規則 A (靜默跳過)：空表不具備導出意義，直接返回原始對象，不存入隊列
            // 对于空数据表采取宽容策略，避免不必要的错误
            Err(XlsxError::EmptyDataFrame) => return Ok(self),
            // 規則 B (自動修復)：名稱非法時，放棄用戶名稱，改用系統預設名稱重試
            // 此時已知 df 不為空，第二次嘗試是安全的
            Err(XlsxError::InvalidName(_)) => {
                // 使用默认名称重新尝试创建工作表
                let fallback_name = get_default_name(self.sheets.len());
                // 再次調用 from_dataframe，此時使用安全名稱（已知 df 不為空，所以這次一定會 Ok）
                WorkSheet::from_dataframe(df, fallback_name, style_map, merge_ranges)?
            }
            // 其他嚴重錯誤：直接包裝並向上拋出
            // 保留原始错误信息，便于问题诊断
            Err(e) => return Err(Box::new(e)),
        };

        // 4. 名稱重複檢查：
        // Excel 不允許同名工作表。這是全局級別的衝突，必須由 Workbook 攔截。
        if self.sheets.iter().any(|s| s.name == task.name) {
            return Err(Box::new(XlsxError::DuplicateName(task.name)));
        }

        // 5. 樣式名存在性檢查：
        // 確保數據在寫入時能找到對應的格式定義，防止 save 時出現懸空引用。
        if let Some(ref map) = task.style_map {
            for style_name in map.values() {
                if !self.styles.contains_key(style_name.as_ref()) {
                    return Err(Box::new(XlsxError::UnknownStyle(style_name.clone())));
                }
            }
        }

        // 通過所有校驗，將任務存入隊列
        self.sheets.push(task);
        Ok(self)
    }

    /// 使用样式工厂向工作簿中添加带样式的工作表
    ///
    /// 便捷方法，用于将 DataFrame 数据结合样式工厂生成的样式映射和合并区域
    /// 一体化地添加到工作簿中。通过自动化样式计算减少用户的配置负担。
    ///
    /// # 参数说明
    ///
    /// * `df`: 待导出的 DataFrame 数据源
    ///   - 提供给样式工厂进行样式计算
    ///   - 作为工作表的数据内容
    /// * `name`: 可选的工作表名称
    ///   - None：使用默认名称生成策略
    ///   - Some(name)：使用提供的名称（会进行验证）
    /// * `style_factory`: 可选的样式工厂引用
    ///   - 负责基于 DataFrame 数据计算样式映射
    ///   - 不会被消费，可重复使用
    /// * `merge_factory`: 可选的合并区域工厂引用
    ///   - 负责基于 DataFrame 数据设置合并区域
    ///   - 不会被消费，可重复使用
    pub fn insert_with_factory(
        self,
        df: DataFrame,
        name: Option<String>,
        style_factory: Option<StyleFactory>,
        merge_factory: Option<MergeFactory>,
    ) -> Result<Self, Box<dyn Error>> {
        // 1. 執行样式工厂引擎：计算该 DataFrame 的物理样式地图
        // 产出类型为 HashMap<(u32, u16), Arc<str>>
        // 此步骤会进行样式规则评估和样式映射生成
        let style_map = match style_factory {
            Some(factory) => Some(factory.execute(&df)?),
            None => None,
        };

        // 2. 執行合并区域工厂引擎：计算该 DataFrame 的合并区域设置
        // 产出类型为 Vec<(u32, u16, u32, u16)>，表示四个角的坐标
        // 此步骤会进行合并区域规则评估和合并区域生成
        let merge_ranges = match merge_factory {
            Some(factory) => Some(factory.execute(&df)?),
            None => None,
        };

        // 3. 调用原有的成品 insert 函数
        // 注意：这里直接将计算好的 style_map 和 merge_ranges 传入
        // 复用现有的工作表添加逻辑，确保行为一致性
        self.insert(df, name, style_map, merge_ranges)
    }

    /// 使用 JSON 配置向工作簿中添加带样式的工作表
    ///
    /// 便捷方法，用于直接从 JSON 配置数据创建样式工厂和合并工厂，并应用到工作表添加过程。
    /// 通过配置驱动的方式实现样式定义与代码的完全分离，支持灵活的样式配置。
    ///
    /// # 参数说明
    ///
    /// * `df`: 待导出的 DataFrame 数据源
    ///   - 提供给样式工厂进行样式规则评估
    ///   - 作为工作表的实际数据内容
    /// * `name`: 可选的工作表名称
    ///   - None：使用系统默认命名策略
    ///   - Some(name)：使用提供的名称（会进行验证和清理）
    /// * `config`: 对 JSON 配置的不可变引用
    ///   - 包含样式规则定义和合并规则的结构化数据
    ///   - 支持 null 值，由工厂内部进行容错处理
    pub fn insert_with_config(
        self,
        df: DataFrame,
        name: Option<String>,
        config: &serde_json::Value,
    ) -> Result<Self, Box<dyn Error>> {
        // 直接傳入引用的片段（可能是 Null），Factory 內部會自我癒合
        let style_factory = match StyleFactory::new(config.clone()) {
            Ok(factory) => Some(factory),
            Err(_) => None,
        };
        let merge_factory = match MergeFactory::new(config.clone()) {
            Ok(factory) => Some(factory),
            Err(_) => None,
        };

        Ok(self.insert_with_factory(df, name, style_factory, merge_factory)?)
    }

    // ------------------------------------------------------------------------
    // 保存方法
    // ------------------------------------------------------------------------

    /// 将工作簿保存为 Excel 文件
    ///
    /// 核心导出方法，负责将所有注册的工作表任务转换为实际的 Excel 文件。
    /// 通过精细的样式处理和类型转换，确保数据在 Excel 中的正确呈现。
    ///
    /// # 参数说明
    ///
    /// * `path`: 目标 Excel 文件的保存路径
    ///   - 可以是相对路径或绝对路径
    ///   - 文件格式根据扩展名自动识别（.xls 或 .xlsx）
    ///   - 如果文件已存在会被覆盖
    pub fn save(&self, path: &str) -> Result<(), Box<dyn Error>> {
        // 检测文件扩展名，.xls 使用 XlsWorkbook，其他使用 xlsx
        let is_xls = path.to_lowercase().ends_with(".xls");

        if is_xls {
            self.save_as_xls(path)
        } else {
            self.save_as_xlsx(path)
        }
    }

    /// 将工作簿保存为 .xls 格式（Excel 97-2003）
    ///
    /// 注意：此过程会丢失所有样式信息（style_map 被忽略）
    ///
    /// # 参数
    /// * `path`: 目标文件路径
    ///
    /// # 返回值
    /// * `Ok(())` - 保存成功
    /// * `Err(Box<dyn Error>)` - 保存失败（空工作簿或写入错误）
    fn save_as_xls(&self, path: &str) -> Result<(), Box<dyn Error>> {
        use cfb::CompoundFile;
        use std::fs::File;
        use std::io::Write;

        // 检查是否有工作表
        if self.sheets.is_empty() {
            return Err("工作簿为空，没有可写入的工作表".into());
        }

        // 生成工作簿 BIFF 数据
        let workbook_stream = self.generate_xls_biff_data()?;

        // 写入 CFB 文件
        let file = File::create(path)?;
        let mut compound_file = CompoundFile::create(file)?;
        let mut stream = compound_file.create_stream("Workbook")?;
        stream.write_all(&workbook_stream)?;
        stream.flush()?;
        drop(stream);
        drop(compound_file);

        Ok(())
    }

    /// 生成 XLS 格式的 BIFF 数据
    fn generate_xls_biff_data(&self) -> Result<Vec<u8>, Box<dyn Error>> {
        use crate::xls_records::*;

        // 第一步：创建共享字符串表
        let mut sst = SharedStringTable::new();

        // 第二步：生成每个工作表的 BIFF 数据
        let sheet_biff_data: Vec<Vec<u8>> = self
            .sheets
            .iter()
            .map(|sheet| sheet.to_biff_data(&mut sst))
            .collect();

        // 如果没有成功生成任何工作表数据，返回错误
        if sheet_biff_data.is_empty() {
            return Err("没有可写入的工作表".into());
        }

        let mut result = Vec::new();

        // 第三步：写入 Workbook Globals 头部记录
        let header_data = self.write_xls_workbook_headers();
        let header_len = header_data.len();
        result.extend_from_slice(&header_data);

        // 第四步：创建 pending BOUNDSHEET 记录并计算大小
        let mut pending_boundsheets: Vec<BoundSheetRecord> = self
            .sheets
            .iter()
            .map(|sheet| BoundSheetRecord::new_pending(&sheet.name))
            .collect();
        let boundsheets_total_len: usize = pending_boundsheets
            .iter()
            .map(|r| r.serialize().len())
            .sum();

        // 第五步：计算 BOUNDSHEET 偏移
        let sst_data = SSTRecord::from(&sst).serialize();
        let eof_data = EofRecord::default().serialize();
        let after_boundsheets_len = sst_data.len() + eof_data.len();

        // 每个 sheet 的偏移 = header_len + boundsheets_total_len + after_boundsheets_len + 前面所有 sheet 数据长度之和
        let mut current_offset =
            (header_len + boundsheets_total_len + after_boundsheets_len) as u32;
        for sheet_data in &sheet_biff_data {
            current_offset += sheet_data.len() as u32;
        }

        // 从后往前设置偏移（因为 current_offset 现在是末尾位置）
        for (i, sheet_data) in sheet_biff_data.iter().enumerate().rev() {
            current_offset -= sheet_data.len() as u32;
            pending_boundsheets[i].set_offset(current_offset);
        }

        // 第六步：写入带正确偏移的 BOUNDSHEET 记录
        for boundsheet in &pending_boundsheets {
            result.extend_from_slice(&boundsheet.serialize());
        }

        // 第七步：写入 SST
        result.extend_from_slice(&sst_data);

        // 第八步：写入 EOF（Workbook Globals 子流结束）
        result.extend_from_slice(&eof_data);

        // 第九步：写入工作表数据
        for sheet_data in sheet_biff_data {
            result.extend_from_slice(&sheet_data);
        }

        Ok(result)
    }

    /// 写入 Workbook Globals 头部记录（从 BOF 到 UseSelfs）
    fn write_xls_workbook_headers(&self) -> Vec<u8> {
        use crate::xls_records::*;
        let mut result = Vec::new();

        // BOF - Workbook globals
        result.extend_from_slice(&BoFRecord::new(BofType::WorkbookGlobals).serialize());

        // InterfaceHdr
        result.extend_from_slice(&InterfaceHdrRecord::default().serialize());

        // MMS
        result.extend_from_slice(&MMSRecord::default().serialize());

        // InterfaceEnd
        result.extend_from_slice(&InterfaceEndRecord::default().serialize());

        // WriteAccess
        result.extend_from_slice(&WriteAccessRecord::new("yujitui").serialize());

        // Codepage
        result.extend_from_slice(&CodepageRecord::new().serialize());

        // DSF
        result.extend_from_slice(&DSFRecord::default().serialize());

        // TabID
        result.extend_from_slice(&TabIDRecord::new(self.sheets.len() as u16).serialize());

        // FnGroupCount
        result.extend_from_slice(&FnGroupCountRecord::default().serialize());

        // Workbook Protection Block
        result.extend_from_slice(&WindowProtectRecord::default().serialize());
        result.extend_from_slice(&ProtectRecord::default().serialize());
        result.extend_from_slice(&ObjectProtectRecord::default().serialize());
        result.extend_from_slice(&PasswordRecord::default().serialize());
        result.extend_from_slice(&Prot4RevRecord::default().serialize());
        result.extend_from_slice(&Prot4RevPassRecord::default().serialize());

        // Backup
        result.extend_from_slice(&BackupRecord::default().serialize());

        // HideObj
        result.extend_from_slice(&HideObjRecord::default().serialize());

        // Window1
        result.extend_from_slice(&Window1Record::default().serialize());

        // DateMode
        result.extend_from_slice(&DateModeRecord::default().serialize());

        // Precision
        result.extend_from_slice(&PrecisionRecord::default().serialize());

        // RefreshAll
        result.extend_from_slice(&RefreshAllRecord::default().serialize());

        // BookBool
        result.extend_from_slice(&BookBoolRecord::default().serialize());

        // Fonts (8 fonts: 4x Arial + 1x Times New Roman + 3x Arial)
        result.extend_from_slice(&self.write_xls_default_fonts());

        // Number Formats
        result.extend_from_slice(&self.write_xls_default_formats());

        // XF records
        result.extend_from_slice(&self.write_xls_default_xf_records());

        // Style
        result.extend_from_slice(&StyleRecord::default().serialize());

        // Palette
        result.extend_from_slice(&PaletteRecord::default().serialize());

        // UseSelfs
        result.extend_from_slice(&UseSelfsRecord::default().serialize());

        result
    }

    /// 写入默认字体记录
    fn write_xls_default_fonts(&self) -> Vec<u8> {
        use crate::xls_records::*;
        let mut result = Vec::new();

        for _ in 0..5 {
            let font = Font::new("Arial");
            result.extend_from_slice(&FontRecord::new(font).serialize());
        }

        let times = Font::new("Times New Roman").with_bold();
        result.extend_from_slice(&FontRecord::new(times).serialize());

        for _ in 0..2 {
            let font = Font::new("Arial");
            result.extend_from_slice(&FontRecord::new(font).serialize());
        }

        result
    }

    /// 写入默认数字格式记录
    fn write_xls_default_formats(&self) -> Vec<u8> {
        use crate::xls_records::*;
        let mut result = Vec::new();

        let general_format = NumberFormatRecord::new(0x00A4, "General");
        result.extend_from_slice(&general_format.serialize());

        result
    }

    /// 写入默认 XF 记录
    fn write_xls_default_xf_records(&self) -> Vec<u8> {
        use crate::xls_records::*;
        let mut result = Vec::new();

        // 使用 16 个默认 XF 记录（简化实现，全部是 Style XF）
        for _ in 0..16 {
            result.extend_from_slice(&XFRecord::default().serialize());
        }

        result
    }

    // ------------------------------------------------------------------------
    // XLSX 格式专用方法
    // ------------------------------------------------------------------------

    /// 将工作簿保存为 .xlsx 格式（Excel 2007+）
    ///
    /// 保留所有样式信息
    ///
    /// # 参数
    /// * `path`: 目标文件路径
    ///
    /// # 返回值
    /// * `Ok(())` - 保存成功
    /// * `Err(Box<dyn Error>)` - 保存失败
    fn save_as_xlsx(&self, path: &str) -> Result<(), Box<dyn Error>> {
        // 创建新的 Excel 工作簿实例，作为导出的目标容器
        let mut workbook = rust_xlsxwriter::Workbook::new();

        // --- 1. 樣式池安全預取：---
        // fallback_fmt 用於解決 Rust 借用檢查器對臨時引用的限制。
        // default_fmt 是你在 new() 中定義的全局 UI 規範（如微軟雅黑+邊框）。
        let fallback_fmt = Format::default(); // 系统回退样式（无样式）
        let default_fmt = self.styles.get("default"); // 全局默认样式（基础格式）

        // 遍历所有注册的工作表任务
        for sheet in &self.sheets {
            // 为当前工作表创建 Excel 工作表实例
            let worksheet = workbook.add_worksheet();
            // 设置工作表名称，支持中文和特殊字符（会进行验证）
            worksheet.set_name(sheet.name.as_str())?;

            // --- 0. 處理 Header 特殊區域 ---
            let mut current_row: u32 = 0;
            for region in &sheet.regions {
                if matches!(region.region_type, crate::sheet_region::RegionType::Header) {
                    current_row = self.write_region_to_worksheet(
                        worksheet,
                        region,
                        current_row,
                        &fallback_fmt,
                    )?;
                }
            }

            // 计算主数据的起始行
            let data_start_row = current_row;

            // 在写入数据前先应用合并区域（注意：合并区域坐标是相对于主数据区域的）
            if let Some(merge_ranges) = &sheet.merge_ranges {
                for merge_range in merge_ranges {
                    worksheet.merge_range(
                        data_start_row + merge_range.0, // first_row（加上偏移）
                        merge_range.1,                  // first_col
                        data_start_row + merge_range.2, // last_row（加上偏移）
                        merge_range.3,                  // last_col
                        "",                             // 数据为空白字符串
                        &fallback_fmt,                  // 使用标准样式保持一致性
                    )?;
                }
            }

            // --- 1. 遍历 cells 写入数据 ---
            // sheet.cells[0] 是表头，sheet.cells[1..] 是数据
            for (row_idx, row) in sheet.cells.iter().enumerate() {
                let r = data_start_row + row_idx as u32;

                for (col_idx, cell_opt) in row.iter().enumerate() {
                    let c = col_idx as u16;

                    // 获取单元格样式
                    let cell_fmt = sheet
                        .style_map
                        .as_ref()
                        .and_then(|m| m.get(&(row_idx as u32, c)))
                        .and_then(|name| self.styles.get(name.as_ref()))
                        .or_else(|| {
                            // 表头行（row_idx == 0）自动应用 header 样式
                            if row_idx == 0 && sheet.style_map.is_none() {
                                self.styles.get("header")
                            } else {
                                None
                            }
                        })
                        .or(default_fmt)
                        .unwrap_or(&fallback_fmt);

                    // 写入单元格
                    if let Some(cell) = cell_opt {
                        use crate::cell::Cell;
                        match cell {
                            Cell::Text(s) => {
                                worksheet.write_with_format(r, c, s.as_str(), cell_fmt)?;
                            }
                            Cell::Number(n) => {
                                worksheet.write_number_with_format(r, c, *n, cell_fmt)?;
                            }
                            Cell::Boolean(b) => {
                                worksheet.write_boolean_with_format(r, c, *b, cell_fmt)?;
                            }
                        }
                    } else {
                        // 空单元格，写入 blank 保持样式一致性
                        worksheet.write_blank(r, c, cell_fmt)?;
                    }
                }
            }

            // --- 5. 處理 Footer 特殊區域 ---
            let mut footer_start_row = data_start_row + sheet.cells.len() as u32;
            for region in &sheet.regions {
                if matches!(region.region_type, crate::sheet_region::RegionType::Footer) {
                    footer_start_row = self.write_region_to_worksheet(
                        worksheet,
                        region,
                        footer_start_row,
                        &fallback_fmt,
                    )?;
                }
            }

            // --- 6. 自動列寬適配：---
            // 必須在數據寫入完成後調用，確保 Excel 引擎能計算出最長內容的佔位。
            worksheet.autofit(); // 自动调整列宽以适应内容
        }

        // 将内存中的工作簿保存到指定路径的文件中
        workbook.save(path)?;
        Ok(()) // 成功完成保存操作
    }
}

impl fmt::Debug for Workbook {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Workbook")
            .field("sheets", &self.sheets) // 只打印實現了 Debug 的 sheets
            .field("styles_count", &self.styles.len()) // 打印樣式數量作為替代
            .finish()
    }
}
