//! Excel 工作表导出任务定义模块
//!
//! 本模块定义了 Excel 工作表导出的核心数据结构，包括：
//! - WorkSheet: 单个工作表的导出任务
//! - 工作表名称验证机制
//! - 数据和样式映射的统一管理

use crate::cell::Cell;
use crate::error::XlsxError;
use crate::sheet_region::SheetRegion;
use crate::xls_records::{
    row_data_to_cell_records, BiffRecord, BoFRecord, BofType, BottomMarginRecord, CalcCountRecord,
    CalcModeRecord, DefaultRowHeightRecord, DeltaRecord, DimensionsRecord, EofRecord, FooterRecord,
    GridSetRecord, GutsRecord, HCenterRecord, HeaderRecord, IterationRecord, LeftMarginRecord,
    PrintGridLinesRecord, PrintHeadersRecord, RefModeRecord, RightMarginRecord, RowRecord,
    ScenProtectRecord, SetupPageRecord, SharedStringTable, TopMarginRecord, VCenterRecord,
    WSBoolRecord, Window2Record, WorksheetObjectProtectRecord, WorksheetProtectRecord,
    WorksheetWindowProtectRecord,
};
use polars::datatypes::AnyValue;
use polars::prelude::{DataFrame, Series};
// Column is re-exported from polars::prelude
use std::collections::HashMap;

/// Excel 工作表名称中禁止出现的特殊字符集合
///
/// 根据 Microsoft Excel 规范，工作表名称不能包含以下字符：
/// - 反斜杠 `\` - 文件系统路径分隔符冲突
/// - 正斜杠 `/` - 文件系统路径分隔符冲突
/// - 问号 `?` - URL 查询参数分隔符冲突
/// - 星号 `*` - 通配符冲突
/// - 冒号 `:` - 驱动器标识符冲突
/// - 左方括号 `[` - 引用语法冲突
/// - 右方括号 `]` - 引用语法冲突
const FORBIDDEN_SHEET_NAME_CHARS: [char; 7] = ['\\', '/', '?', '*', ':', '[', ']'];

/// Excel 工作表导出任务描述符
///
/// 封装了单个工作表导出所需的所有信息。
/// 使用统一的 cells 数据存储，不再依赖 DataFrame。
#[derive(Debug)]
pub struct WorkSheet {
    /// 工作表显示名称
    ///
    /// 在 Excel 中显示的工作表标签名称，需符合命名规范。
    pub name: String,

    /// 单元格数据
    ///
    /// 二维向量存储所有单元格数据。
    /// 第 0 行是表头（列名），第 1 行开始是数据。
    /// `None` 表示空单元格。
    pub cells: Vec<Vec<Option<Cell>>>,

    /// 列名列表
    ///
    /// 缓存列名，避免重复从 cells[0] 提取。
    pub column_names: Vec<String>,

    /// 单元格样式映射表（可选）
    ///
    /// 提供精确到单元格级别的样式控制能力，键为 (行, 列) 坐标，
    /// 值为预定义样式池中的样式名称。
    pub style_map: Option<HashMap<(u32, u16), std::sync::Arc<str>>>,

    /// 合并单元格区域列表（可选）
    ///
    /// 定义需要在 Excel 中合并的单元格区域，每个元组表示一个合并区域：
    /// `(起始行, 起始列, 结束行, 结束列)`。
    pub merge_ranges: Option<Vec<(u32, u16, u32, u16)>>,

    /// 特殊区域列表（表头、脚注等）
    ///
    /// 存储独立于主数据区域的特殊区域，按添加顺序写入。
    /// Header 类型的区域显示在主数据上方，
    /// Footer 类型的区域显示在主数据下方。
    pub regions: Vec<SheetRegion>,
}

impl WorkSheet {
    /// 主构造函数：从 cells 创建 WorkSheet
    ///
    /// 这是最通用的创建方式，xlsx/xls 读取都使用这个。
    ///
    /// # 参数
    /// * `cells` - 单元格数据（第 0 行必须是表头）
    /// * `name` - 工作表名称
    /// * `style_map` - 可选的单元格样式映射
    /// * `merge_ranges` - 可选的合并单元格区域列表
    ///
    /// # 返回值
    /// * `Ok(WorkSheet)` - 成功创建
    /// * `Err(XlsxError)` - 创建失败
    pub fn new(
        cells: Vec<Vec<Option<Cell>>>,
        name: String,
        style_map: Option<HashMap<(u32, u16), std::sync::Arc<str>>>,
        merge_ranges: Option<Vec<(u32, u16, u32, u16)>>,
    ) -> Result<Self, XlsxError> {
        // 1. 验证 cells
        if cells.is_empty() || cells[0].is_empty() {
            return Err(XlsxError::EmptyDataFrame);
        }

        // 2. 名称规范性校验
        let trimmed = name.trim();
        if trimmed.is_empty()
            || trimmed.chars().count() > 31
            || trimmed.contains(&FORBIDDEN_SHEET_NAME_CHARS[..])
        {
            return Err(XlsxError::InvalidName(name));
        }

        // 3. 从第 0 行提取列名
        let column_names: Vec<String> = cells[0]
            .iter()
            .map(|cell| match cell {
                Some(Cell::Text(s)) => s.clone(),
                _ => "Column".to_string(),
            })
            .collect();

        // 4. 样式坐标清洗
        let style_map = style_map.and_then(|mut map| {
            let max_row = cells.len() as u32;
            let max_col = column_names.len() as u16;

            map.retain(|&(r, c), style_name| {
                r < max_row && c < max_col && !style_name.trim().is_empty()
            });

            if map.is_empty() {
                None
            } else {
                Some(map)
            }
        });

        // 5. 合并区域清洗
        let merge_ranges = merge_ranges.and_then(|mut ranges| {
            let max_row = cells.len() as u32;
            let max_col = column_names.len() as u16;

            ranges.retain(|&(start_r, start_c, end_r, end_c)| {
                start_r <= end_r && start_c <= end_c && end_r < max_row && end_c < max_col
            });

            if ranges.is_empty() {
                None
            } else {
                Some(ranges)
            }
        });

        Ok(WorkSheet {
            name,
            cells,
            column_names,
            style_map,
            merge_ranges,
            regions: Vec::new(),
        })
    }

    /// 从 DataFrame 创建 WorkSheet
    ///
    /// 便利函数：将 DataFrame 转换为 cells，然后调用 `Self::new()`。
    ///
    /// # 参数
    /// * `df`: 待导出的 Polars DataFrame 数据源
    /// * `name`: 工作表显示名称
    /// * `style_map`: 可选的单元格样式映射
    /// * `merge_ranges`: 可选的合并单元格区域列表
    ///
    /// # 返回值
    /// * `Ok(WorkSheet)` - 成功创建的工作表实例
    /// * `Err(XlsxError)` - 创建过程中发生的验证错误
    pub fn from_dataframe(
        df: DataFrame,
        name: String,
        style_map: Option<HashMap<(u32, u16), std::sync::Arc<str>>>,
        merge_ranges: Option<Vec<(u32, u16, u32, u16)>>,
    ) -> Result<Self, XlsxError> {
        // 1. 数据校验：拒绝空表
        if df.height() == 0 || df.width() == 0 {
            return Err(XlsxError::EmptyDataFrame);
        }

        // 2. 提取列名
        let column_names: Vec<String> = df.columns().iter().map(|c| c.name().to_string()).collect();

        // 3. 转换 DataFrame 为 cells
        let mut cells: Vec<Vec<Option<Cell>>> = Vec::new();

        // 第 0 行：表头
        cells.push(
            column_names
                .iter()
                .map(|name| Some(Cell::Text(name.clone())))
                .collect(),
        );

        // 数据行
        for row_idx in 0..df.height() {
            let mut row: Vec<Option<Cell>> = Vec::new();
            for col in df.columns() {
                let val = col.get(row_idx).map_err(|e| {
                    XlsxError::GenericError(format!(
                        "Failed to get value at ({}, {:?}): {}",
                        row_idx,
                        col.name(),
                        e
                    ))
                })?;
                row.push(Cell::from_any_value(&val));
            }
            cells.push(row);
        }

        // 4. 调用主构造函数
        Self::new(cells, name, style_map, merge_ranges)
    }

    /// 转换为 DataFrame（供外部分析使用）
    ///
    /// # 返回值
    /// * `Ok(DataFrame)` - 成功转换
    /// * `Err(XlsxError)` - 转换失败
    pub fn to_dataframe(&self) -> Result<DataFrame, XlsxError> {
        if self.cells.len() <= 1 {
            // 只有表头，没有数据
            return Err(XlsxError::EmptyDataFrame);
        }

        let mut series_vec: Vec<Series> = Vec::new();

        for (col_idx, col_name) in self.column_names.iter().enumerate() {
            let mut data: Vec<AnyValue> = Vec::new();

            // 从第 1 行开始（跳过表头）
            for row_idx in 1..self.cells.len() {
                let cell = self.cells[row_idx]
                    .get(col_idx)
                    .and_then(|opt| opt.as_ref());
                let any_val = match cell {
                    Some(Cell::Text(s)) => AnyValue::StringOwned(s.clone().into()),
                    Some(Cell::Number(n)) => AnyValue::Float64(*n),
                    Some(Cell::Boolean(b)) => AnyValue::Boolean(*b),
                    None => AnyValue::Null,
                };
                data.push(any_val);
            }

            let series = Series::from_any_values(col_name.as_str().into(), &data, true)
                .map_err(|e| XlsxError::GenericError(format!("创建 Series 失败: {}", e)))?;
            series_vec.push(series);
        }

        // 将 Series 转换为 Columns
        let columns: Vec<polars::prelude::Column> =
            series_vec.into_iter().map(|s| s.into()).collect();
        DataFrame::new(columns[0].len(), columns)
            .map_err(|e| XlsxError::GenericError(format!("创建 DataFrame 失败: {}", e)))
    }

    /// 获取数据行数（不含表头）
    pub fn data_row_count(&self) -> usize {
        self.cells.len().saturating_sub(1)
    }

    /// 获取列数
    pub fn col_count(&self) -> usize {
        self.column_names.len()
    }

    /// 添加特殊区域到工作表
    pub fn add_region(&mut self, region: SheetRegion) {
        self.regions.push(region);
    }

    /// 获取所有 Header 类型的区域
    pub fn header_regions(&self) -> Vec<&SheetRegion> {
        self.regions
            .iter()
            .filter(|r| matches!(r.region_type, crate::sheet_region::RegionType::Header))
            .collect()
    }

    /// 获取所有 Footer 类型的区域
    pub fn footer_regions(&self) -> Vec<&SheetRegion> {
        self.regions
            .iter()
            .filter(|r| matches!(r.region_type, crate::sheet_region::RegionType::Footer))
            .collect()
    }

    /// 生成 BIFF8 数据（用于 .xls 写入）
    ///
    /// 此方法将工作表序列化为 BIFF8 格式的字节流。
    /// 从 XlsSheet::get_biff_data 平移。
    ///
    /// # 参数
    /// * `sst` - 共享字符串表，用于存储文本单元格
    ///
    /// # 返回值
    /// BIFF8 格式的字节流
    pub fn to_biff_data(&self, sst: &mut SharedStringTable) -> Vec<u8> {
        let mut result = Vec::new();

        result.extend_from_slice(&BoFRecord::new(BofType::Worksheet).serialize());

        result.extend_from_slice(&CalcModeRecord::default().serialize());
        result.extend_from_slice(&CalcCountRecord::default().serialize());
        result.extend_from_slice(&RefModeRecord::default().serialize());
        result.extend_from_slice(&DeltaRecord::default().serialize());
        result.extend_from_slice(&IterationRecord::default().serialize());

        result.extend_from_slice(&GutsRecord::default().serialize());
        result.extend_from_slice(&DefaultRowHeightRecord::default().serialize());
        result.extend_from_slice(&WSBoolRecord::default().serialize());
        result.extend_from_slice(&DimensionsRecord::from(&self.cells).serialize());

        result.extend_from_slice(&PrintHeadersRecord::default().serialize());
        result.extend_from_slice(&PrintGridLinesRecord::default().serialize());
        result.extend_from_slice(&GridSetRecord::default().serialize());
        result.extend_from_slice(&HeaderRecord::default().serialize());
        result.extend_from_slice(&FooterRecord::default().serialize());
        result.extend_from_slice(&HCenterRecord::default().serialize());
        result.extend_from_slice(&VCenterRecord::default().serialize());
        result.extend_from_slice(&LeftMarginRecord::default().serialize());
        result.extend_from_slice(&RightMarginRecord::default().serialize());
        result.extend_from_slice(&TopMarginRecord::default().serialize());
        result.extend_from_slice(&BottomMarginRecord::default().serialize());
        result.extend_from_slice(&SetupPageRecord::default().serialize());

        result.extend_from_slice(&WorksheetProtectRecord::default().serialize());
        result.extend_from_slice(&WorksheetWindowProtectRecord::default().serialize());
        result.extend_from_slice(&ScenProtectRecord::default().serialize());
        result.extend_from_slice(&WorksheetObjectProtectRecord::default().serialize());

        for (row_idx, row) in self.cells.iter().enumerate() {
            if row.iter().any(|c| c.is_some()) {
                result.extend_from_slice(&RowRecord::from_row_data(row_idx, row).serialize());
                result.extend_from_slice(&row_data_to_cell_records(row_idx, row, 0, sst));
            }
        }

        result.extend_from_slice(&Window2Record::default().serialize());
        result.extend_from_slice(&EofRecord::default().serialize());

        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use polars::prelude::NamedFrom;

    #[test]
    fn test_worksheet_from_dataframe() {
        use polars::frame::column::Column;
        let columns: Vec<Column> = vec![
            Series::new("Name".into(), vec!["Alice", "Bob"]).into(),
            Series::new("Age".into(), vec![30_i64, 25]).into(),
        ];
        let df = DataFrame::new(2, columns).unwrap();

        let sheet = WorkSheet::from_dataframe(df, "Test".to_string(), None, None).unwrap();

        assert_eq!(sheet.name, "Test");
        assert_eq!(sheet.data_row_count(), 2);
        assert_eq!(sheet.col_count(), 2);
        assert_eq!(sheet.column_names, vec!["Name", "Age"]);
    }

    #[test]
    fn test_worksheet_to_dataframe() {
        let cells = vec![
            vec![
                Some(Cell::Text("Name".to_string())),
                Some(Cell::Text("Age".to_string())),
            ],
            vec![
                Some(Cell::Text("Alice".to_string())),
                Some(Cell::Number(30.0)),
            ],
            vec![
                Some(Cell::Text("Bob".to_string())),
                Some(Cell::Number(25.0)),
            ],
        ];

        let sheet = WorkSheet::new(cells, "Test".to_string(), None, None).unwrap();
        let df = sheet.to_dataframe().unwrap();

        assert_eq!(df.height(), 2);
        assert_eq!(df.width(), 2);
    }
}
