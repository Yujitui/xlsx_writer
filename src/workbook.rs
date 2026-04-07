//! Excel 工作簿导出管理器模块
//!
//! 本模块定义了 Excel 导出的核心容器结构 Workbook，负责：
//! - 管理全局样式配置（样式池）
//! - 维护工作表导出任务队列
//! - 协调多工作表的统一导出流程
use crate::error::XlsxError;
use crate::merge_factory::MergeFactory;
use crate::style_factory::StyleFactory;
use crate::style_library::StyleLibrary;
use crate::worksheet::WorkSheet;
use calamine::{open_workbook, Data, DataType, Reader, Xlsx};
use chrono::NaiveDate;
use polars::prelude::*;
use rust_xlsxwriter::*;
use std::collections::{HashMap};
use std::error::Error;
use std::fmt;
use std::io::BufReader;
use std::path::Path;
use crate::prelude::ReadSheet;

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

impl Workbook {
    /// 获取工作簿中工作表的数量
    ///
    /// 用于查询当前工作簿中已添加的工作表总数，
    /// 便于了解工作簿的规模和进行相关统计。
    ///
    /// # 返回值
    ///
    /// 返回 `usize` 类型的工作表数量：
    /// - `0` 表示空工作簿（尚未添加任何工作表）
    /// - `n` 表示包含 n 个工作表
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// 按索引获取工作表名称
    ///
    /// # 参数
    /// * `index` - 工作表索引
    ///
    /// # 返回值
    /// * `Some(&str)` - 工作表名称
    /// * `None` - 索引无效
    pub fn get_sheet_name(&self, index: usize) -> Option<&str> {
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
    pub fn get_sheet_by_name(&self, name: &str) -> Option<&WorkSheet> {
        self.sheets.iter().find(|sheet| sheet.name == name)
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
    pub fn get_sheet_style_map_simple(&self, index: usize) -> Option<&HashMap<(u32, u16), Arc<str>>> {
        self.sheets.get(index).and_then(|sheet| sheet.style_map.as_ref())
    }

}

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
            "PingFang SC"      // macOS: 苹方
        } else {
            "sans-serif"       // Linux 或其他: 通用無襯線字體
        };

        // 2. 定義「標準樣式」：
        // 這是報表中絕大多數單元格的底層模板，包含了邊框、字體和對齊方式。
        let standard_fmt = Format::new()
            .set_font_name(font_name) // 設置微軟雅黑
            .set_font_size(11)                 // 設置常用字號
            .set_border(FormatBorder::Thin)   // 設置細邊框（四周）
            .set_align(FormatAlign::Center)   // 水平居中
            .set_align(FormatAlign::VerticalCenter); // 垂直居中

        // 3. 快速派生「表頭樣式」：
        // 通過 clone() 繼承 standard_fmt 的屬性，僅對差異項（加粗、背景）進行覆蓋。
        // 這種鏈式調用確保了表頭與數據行在字體、邊框寬度上完全對齊。
        let header_fmt = standard_fmt.clone()
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

    /// 将DataFrame中的指定列转换为字符串类型
    ///
    /// 此函数用于确保某些列始终以字符串形式存储，即使它们在Excel中可能是数值类型。
    /// 这对于处理ID号、邮政编码等需要保持前导零的字段特别有用。
    ///
    /// # 参数
    /// * `df` - 需要处理的DataFrame引用
    /// * `string_columns` - 包含需要转换为字符串的列名向量
    ///
    /// # 返回值
    /// * `Result<DataFrame, XlsxError>` - 成功时返回处理后的DataFrame，失败时返回错误
    fn sanitize_to_string_column(name: &str, data: &[AnyValue]) -> Result<Series, Box<dyn Error>> {
        let sanitized: Vec<AnyValue> = data
            .iter()
            .map(|val| match val {
                // 将数字、布尔等统统转为字符串
                AnyValue::Float64(f) => {
                    // 如果是整数，去掉小数点；否则保留
                    let s = if f.fract() == 0.0 { format!("{:.0}", f) } else { f.to_string() };
                    AnyValue::StringOwned(s.into())
                },
                AnyValue::Int64(i) => AnyValue::StringOwned(i.to_string().into()),
                AnyValue::Boolean(b) => AnyValue::StringOwned(b.to_string().into()),
                // 如果已经是字符串或 Null，保持原样
                AnyValue::String(s) => AnyValue::StringOwned((*s).into()),
                AnyValue::StringOwned(s) => AnyValue::StringOwned(s.clone()),
                AnyValue::Date(d) => {
                    // 假设你引用了 chrono，或者简单地将其转为字符串
                    // 这里的 d 是自 1970-01-01 以来的天数
                    AnyValue::StringOwned(format!("{}", d).into()) // 或者调用专门的日期格式化函数
                },
                AnyValue::Null => AnyValue::Null,
                _ => AnyValue::StringOwned(format!("{:?}", val).into()), // 兜底处理
            })
            .collect();

        Ok(Series::from_any_values(name.into(), &sanitized, true)?)
    }

    /// 从Excel文件中读取数据并转换为Polars DataFrame
    ///
    /// 此函数支持读取xlsx和xlsm格式的Excel文件，能够处理多种数据类型并提供
    /// 灵活的列类型控制选项。自动处理日期、数值、字符串等不同类型的数据。
    ///
    /// # 参数
    /// * `path` - Excel文件的完整路径
    /// * `read_sheet` - 可选的工作表读取配置，如果为None则读取第一个可见工作表
    ///   - `sheet_name`: 指定要读取的工作表名称
    ///   - `force_string_cols`: 可选的列名列表，指定这些列强制作为字符串类型读取
    ///   - `skip_rows`: 可选的行数，指定要跳过的起始行数
    ///
    /// # 返回值
    /// * `Result<DataFrame, XlsxError>` - 成功时返回包含数据的DataFrame，失败时返回错误
    fn read_xlsx(
        excel: &mut Xlsx<BufReader<std::fs::File>>,
        read_sheet: Option<ReadSheet>,
    ) -> Result<DataFrame, XlsxError> {
        // Step 1: 打开 Excel 文件
        // let mut excel: Xlsx<_> =
        //     open_workbook(Path::new(path)).map_err(|e| XlsxError::CalamineError(e))?;

        // Step 2: 获取可用的工作表名列表
        let sheet_names = excel.sheet_names();
        if sheet_names.is_empty() {
            return Err(XlsxError::NoSheetsFound);
        }

        // Step 3: 确定要使用的参数
        let (target_sheet, skip_rows, force_string_cols) = match &read_sheet {
            Some(sheet) => {
                // 验证 sheet 名称是否存在
                if !sheet_names.contains(&sheet.sheet_name) {
                    return Err(XlsxError::SheetNotFound(sheet.sheet_name.clone()));
                }

                let skip_rows = sheet.skip_rows.unwrap_or(0);
                let force_string_cols = sheet.force_string_cols.clone().unwrap_or_default();

                (&sheet.sheet_name, skip_rows, force_string_cols)
            }
            None => {
                // 使用第一个 sheet，使用默认值
                let first_sheet = sheet_names.first().ok_or(XlsxError::NoSheetsFound)?;
                (first_sheet, 0, Vec::new())
            }
        };

        // Step 4: 获取指定工作表的数据范围 (Range)
        let range = excel
            .worksheet_range(target_sheet)
            .map_err(|e| XlsxError::CalamineError(e))?;

        // Step 5: 获取所有数据
        let mut rows = range.rows().skip(skip_rows);

        // Step 6: 提取表头并处理空值/缺失项
        let headers: Vec<String> = {
            let first_row = rows.next();
            first_row
                .ok_or(XlsxError::MissingHeaderRow)?
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

        // Step 7: 初始化每一列的容器
        let mut columns_data: Vec<Vec<AnyValue>> = vec![Vec::new(); headers.len()];

        // Step 8: 从第二行开始遍历数据行
        for row in rows {
            for (col_idx, cell) in row.iter().enumerate() {
                let value = if force_string_cols.contains(&headers[col_idx]) {
                    AnyValue::StringOwned(cell.to_string().into())
                } else {
                    match cell {
                        Data::Int(val) => AnyValue::Int64(*val),
                        Data::Float(val) => AnyValue::Float64(*val),
                        Data::String(val) => AnyValue::String(val),
                        Data::Bool(val) => AnyValue::Boolean(*val),
                        Data::Empty => AnyValue::Null,
                        _ => {
                            if let Some(date) = cell.as_date() {
                                let days = (date - NaiveDate::from_ymd_opt(1970, 1, 1).unwrap()).num_days();
                                AnyValue::Date(days as i32)
                            } else {
                                AnyValue::Null
                            }
                        }
                    }
                };
                columns_data[col_idx].push(value);
            }
        }

        // Step 8: 转换为 Vec<Series>
        let series_vec: Vec<Series> = headers
            .iter()
            .enumerate()
            .map(|(i, name)| {
                let col_name = name.as_str();
                let data = &columns_data[i];

                Series::from_any_values(col_name.into(), data, true).or_else(|_| {
                    eprintln!("警告：列 [{col_name}] 类型不一致，已强制转为 String");
                    Self::sanitize_to_string_column(col_name, data)
                        .map_err(|_| XlsxError::ConversionFailed {
                            column: col_name.to_string(),
                        })
                })
            })
            .collect::<Result<Vec<_>, _>>()?;

        // Step 9: 构造 DataFrame 并返回
        let df = DataFrame::new(series_vec[0].len(), series_vec.into_iter().map(Column::from).collect())
            .map_err(|e| XlsxError::ConversionFailed {
                column: format!("构造 DataFrame 失败: {}", e),
            })?;

        Ok(df)
    }

    /// 从Excel文件读取多个工作表数据
    ///
    /// # 参数
    /// * `path` - Excel文件的完整路径
    /// * `read_sheets` - 工作表读取配置列表，如果为空则读取所有工作表
    ///
    /// # 返回值
    /// 返回当前Workbook实例，便于链式调用
    pub fn read(mut self, path: &str, read_sheets: Vec<ReadSheet>) -> Self {
        let mut excel: Xlsx<BufReader<std::fs::File>> = match open_workbook(Path::new(path)).map_err(XlsxError::CalamineError) {
            Ok(excel) => excel,
            Err(e) => {
                eprintln!("警告: 无法打开Excel文件 '{}': {}", path, e);
                return self
            }
        };

        let read_sheets = if read_sheets.is_empty() {
            let sheet_names = excel.sheet_names();
            if sheet_names.is_empty() {
                eprintln!("警告: Excel文件中没有工作表");
                return self;
            }
            sheet_names.iter().map(|name| ReadSheet::new(name.to_string())).collect()
        }
        else {
            //去除可能重复的sheet_name
            // 去除重复的sheet_name，保持原始顺序
            let mut seen = std::collections::HashSet::new();
            read_sheets
                .into_iter()
                .filter(|sheet| seen.insert(sheet.clone()))
                .collect::<Vec<ReadSheet>>()
        };

        // sheet为空则直接返回self
        if read_sheets.is_empty() { return self };

        for sheet in &read_sheets {
            match Self::read_xlsx(&mut excel, Some(sheet.clone())) {
                Ok(df) => {
                    let worksheet = match WorkSheet::new(df, sheet.sheet_name.clone(), None, None) {
                        Ok(worksheet) => worksheet,
                        Err(e) => {
                            eprintln!("创建工作表 '{}' 失败: {}", sheet.sheet_name, e);
                            continue;
                        }
                    };
                    if let None = self.get_sheet_by_name(worksheet.name.as_str()) {
                        self.sheets.push(worksheet);
                    }
                    else {
                        eprintln!("sheet '{}' 与现有表格重复", sheet.sheet_name);
                        continue;
                    }
                }
                Err(e) => {
                    eprintln!("读取sheet '{}' 失败: {}", sheet.sheet_name, e);
                    continue;
                }
            }
        }
        self
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
    pub fn insert(mut self, df: DataFrame, name: Option<String>,
                  style_map: Option<HashMap<(u32, u16), Arc<str>>>,
                  merge_ranges: Option<Vec<(u32, u16, u32, u16)>>
    ) -> Result<Self, Box<dyn Error>> {
        // 1. 定義輔助閉包：封裝默認命名邏輯，確保命名的一致性與唯一性起點
        let get_default_name = |sheets_len: usize| format!("Sheet {}", sheets_len + 1);

        // 2. 初步確定名稱：優先使用用戶提供，否則生成默認名
        let final_name = name.unwrap_or_else(|| get_default_name(self.sheets.len()));

        // 3. 嘗試構建 WorkSheet 任務：
        // 這裡使用 match 進行細分錯誤處理。注意 df.clone() 是淺拷貝，開銷極低。
        // 這裡 clone() 是必須的，因為如果 InvalidName 發生，我們需要原始數據進行第二次嘗試。
        let task = match WorkSheet::new(df.clone(), final_name.clone(), style_map.clone(), merge_ranges.clone()) {
            Ok(t) => t,
            // 規則 A (靜默跳過)：空表不具備導出意義，直接返回原始對象，不存入隊列
            // 对于空数据表采取宽容策略，避免不必要的错误
            Err(XlsxError::EmptyDataFrame) => return Ok(self),
            // 規則 B (自動修復)：名稱非法時，放棄用戶名稱，改用系統預設名稱重試
            // 此時已知 df 不為空，第二次 new 操作是安全的
            Err(XlsxError::InvalidName(_)) => {
                // 使用默认名称重新尝试创建工作表
                let fallback_name = get_default_name(self.sheets.len());
                // 再次調用 new，此時使用安全名稱（已知 df 不為空，所以這次一定會 Ok）
                WorkSheet::new(df, fallback_name, style_map, merge_ranges)?
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
        merge_factory: Option<MergeFactory>
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
    pub fn insert_with_config(self, df: DataFrame, name: Option<String>, config: &serde_json::Value) -> Result<Self, Box<dyn Error>> {
        // 直接傳入引用的片段（可能是 Null），Factory 內部會自我癒合
        let style_factory = match StyleFactory::new(config.clone()) {
            Ok(factory) => Some(factory),
            Err(_) => None
        };
        let merge_factory = match MergeFactory::new(config.clone()) {
            Ok(factory) => Some(factory),
            Err(_) => None
        };

        Ok(self.insert_with_factory(df, name, style_factory, merge_factory)?)
    }

    /// 将工作簿保存为 Excel 文件
    ///
    /// 核心导出方法，负责将所有注册的工作表任务转换为实际的 Excel 文件。
    /// 通过精细的样式处理和类型转换，确保数据在 Excel 中的正确呈现。
    ///
    /// # 参数说明
    ///
    /// * `path`: 目标 Excel 文件的保存路径
    ///   - 可以是相对路径或绝对路径
    ///   - 文件格式自动识别为 .xlsx
    ///   - 如果文件已存在会被覆盖
    pub fn save(&self, path: &str) -> Result<(), Box<dyn Error>> {
        // 创建新的 Excel 工作簿实例，作为导出的目标容器
        let mut workbook = rust_xlsxwriter::Workbook::new();

        // --- 1. 樣式池安全預取：---
        // fallback_fmt 用於解決 Rust 借用檢查器對臨時引用的限制。
        // default_fmt 是你在 new() 中定義的全局 UI 規範（如微軟雅黑+邊框）。
        let fallback_fmt = Format::default();           // 系统回退样式（无样式）
        let default_fmt = self.styles.get("default");   // 全局默认样式（基础格式）

        // 遍历所有注册的工作表任务
        for sheet in &self.sheets {
            // 为当前工作表创建 Excel 工作表实例
            let worksheet = workbook.add_worksheet();
            // 设置工作表名称，支持中文和特殊字符（会进行验证）
            worksheet.set_name(sheet.name.as_str())?;

            // 在写入数据前先应用合并区域
            if let Some(merge_ranges) = &sheet.merge_ranges {
                for merge_range in merge_ranges {
                    // merge_range 包含四个元素：起始行、起始列、结束行、结束列
                    // 第五、六个参数填写空白以获得这些空白单元格
                    worksheet.merge_range(
                        merge_range.0,  // first_row
                        merge_range.1,  // first_col
                        merge_range.2,  // last_row
                        merge_range.3,  // last_col
                        "",              // 数据为空白字符串
                        &fallback_fmt,  // 使用标准样式保持一致性
                    )?;
                }
            }

            // 獲取 Polars 列引用，准备数据写入
            let columns = sheet.df.columns();
            // 遍历每一列进行数据处理
            for (col_idx, column) in columns.iter().enumerate() {
                let c = col_idx as u16;

                // --- 2. 處理表頭 (Excel Row 0) ---
                // 邏輯升級：
                // 1. 優先查找 style_map 中的定義。
                // 2. 如果 style_map 是 None，則自動應用預設的 "header" 樣式。
                // 3. 如果以上皆否，則使用全局 "default"。
                // 4. 最後使用系統 "fallback"。
                // 四级样式查找机制确保表头样式的正确应用
                let header_cell_fmt = sheet.style_map.as_ref()
                    .and_then(|m| m.get(&(0, c)))         // 嘗試從地圖找特定单元格样式
                    .and_then(|name| self.styles.get(name.as_ref()))  // 查找样式池中的实际样式
                    .or_else(|| {                        // 如果地圖沒定義或地圖不存在
                        if sheet.style_map.is_none() {
                            self.styles.get("header")    // 自動應用預設 header 样式
                        } else {
                            None
                        }
                    })
                    .or(default_fmt)                     // 全局保底样式
                    .unwrap_or(&fallback_fmt);           // 系統保底样式
                // 写入表头单元格，应用确定的样式
                worksheet.write_with_format(0, c, column.name().as_str(), header_cell_fmt)?;

                // --- 3. 處理數據行 (Excel Row 1..N) ---
                // 遍历当前列的所有数据行
                for row_idx in 0..sheet.df.height() {
                    let val = column.get(row_idx)?; // 获取 AnyValue
                    let r = (row_idx + 1) as u32;   // Excel 行索引（跳过表头）

                    // 獲取該單元格專屬樣式或全局保底樣式
                    // 使用与表头相同的四级查找机制
                    let cell_fmt = sheet.style_map.as_ref()
                        .and_then(|m| m.get(&(r, c)))     // 查找特定单元格样式
                        .and_then(|name| self.styles.get(name.as_ref()))  // 获取实际样式
                        .or(default_fmt)                 // 全局默认样式
                        .unwrap_or(&fallback_fmt);       // 系统回退样式

                    // --- 4. 類型分派：將 Polars 豐富的數值類型統化為 Excel 的 f64。 ---
                    // 使用 _with_format 系列方法確保樣式（邊框、字體）被正確應用。
                    match val {
                        // 使用 _with_format 系列方法，傳入 cell_fmt (&Format)
                        // 整数类型转换为 f64 写入
                        AnyValue::Int8(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Int16(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Int32(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Int64(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::UInt32(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        // 浮点数类型处理
                        AnyValue::Float32(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Float64(v) => worksheet.write_number_with_format(r, c, v, cell_fmt)?,
                        // 字符串类型直接写入
                        AnyValue::String(s) => worksheet.write_string_with_format(r, c, s, cell_fmt)?,
                        // 布尔类型处理
                        AnyValue::Boolean(v) => worksheet.write_boolean_with_format(r, c, v, cell_fmt)?,
                        // 日期类型转换（天数偏移）
                        AnyValue::Date(days) => {
                            // 1. 转换基准：Polars 天数 + 25569 = Excel 天数
                            let excel_date_num = (days as f64) + 25569.0;
                            // 2. 在 Excel 中，日期本质上就是带格式的数字
                            worksheet.write_number_with_format(r, c, excel_date_num, cell_fmt)?
                        }
                        // 日期时间类型转换（Unix 时间戳转换）
                        AnyValue::Datetime(v, unit, _) => {
                            // 根据时间单位转换为秒数
                            let seconds = match unit {
                                TimeUnit::Milliseconds => v / 1_000,
                                TimeUnit::Microseconds => v / 1_000_000,
                                TimeUnit::Nanoseconds => v / 1_000_000_000,
                            } as f64;
                            // 将 Unix 秒数转换为 Excel 天数：秒数 / 86400 + 25569
                            let excel_dt_num = (seconds / 86400.0) + 25569.0;

                            worksheet.write_number_with_format(r, c, excel_dt_num, cell_fmt)?
                        }
                        // 空值处理（保持样式一致性）
                        AnyValue::Null => {
                            // Null 值也寫入 blank 以保持單元格邊框一致
                            worksheet.write_blank(r, c, cell_fmt)?
                        },
                        // 其他类型转换为字符串处理
                        _ => {
                            // 處理日期等類型，先轉為字符串
                            let s = format!("{}", val);
                            worksheet.write_string_with_format(r, c, &s, cell_fmt)?
                        },
                    };
                }
            }

            // --- 5. 自動列寬適配：---
            // 必須在數據寫入完成後調用，確保 Excel 引擎能計算出最長內容的佔位。
            worksheet.autofit();  // 自动调整列宽以适应内容
        }

        // 将内存中的工作簿保存到指定路径的文件中
        workbook.save(path)?;
        Ok(())  // 成功完成保存操作
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
