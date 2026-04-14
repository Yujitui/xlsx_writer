//! Excel 工作表导出任务定义模块
//!
//! 本模块定义了 Excel 工作表导出的核心数据结构，包括：
//! - WorkSheet: 单个工作表的导出任务，由多个 SheetRegion 组成
//! - 工作表名称验证机制

use crate::error::XlsxError;
use crate::region_styles::RegionStyles;
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
use polars::prelude::DataFrame;
use std::collections::{HashSet};

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
/// 工作表由多个 SheetRegion 组成，按顺序写入。
#[derive(Debug)]
pub struct WorkSheet {
    /// 工作表显示名称
    ///
    /// 在 Excel 中显示的工作表标签名称，需符合命名规范。
    pub name: String,

    /// 区域列表
    ///
    /// 存储工作表中的所有数据区域，按添加顺序写入。
    /// 每个区域有唯一的名称用于标识。
    pub regions: Vec<SheetRegion>,
}

impl WorkSheet {
    /// 创建新的工作表
    ///
    /// # 参数
    /// * `name` - 工作表名称
    /// * `regions` - 区域列表
    ///
    /// # 返回值
    /// * `Ok(WorkSheet)` - 成功创建
    /// * `Err(XlsxError)` - 创建失败（名称无效或区域名称重复）
    pub fn new(name: impl Into<String>, regions: Vec<SheetRegion>) -> Result<Self, XlsxError> {
        let name = name.into();

        // 1. 验证工作表名称
        let trimmed = name.trim();
        if trimmed.is_empty()
            || trimmed.chars().count() > 31
            || trimmed.contains(&FORBIDDEN_SHEET_NAME_CHARS[..])
        {
            return Err(XlsxError::InvalidName(name));
        }

        // 2. 验证区域名称唯一性
        let mut name_set = HashSet::new();
        for region in &regions {
            if !name_set.insert(region.name.clone()) {
                return Err(XlsxError::GenericError(format!(
                    "Duplicate region name: {}",
                    region.name
                )));
            }
        }

        Ok(WorkSheet { name, regions })
    }

    /// 从 DataFrame 创建 WorkSheet
    ///
    /// 创建一个包含单个区域（默认名"data"）的工作表。
    ///
    /// # 参数
    /// * `df` - Polars DataFrame 数据源
    /// * `sheet_name` - 工作表显示名称
    /// * `region_name` - 区域名称（默认为"data"）
    /// * `style_map` - 可选的单元格样式映射
    /// * `merge_ranges` - 可选的合并单元格区域列表
    ///
    /// # 返回值
    /// * `Ok(WorkSheet)` - 成功创建的工作表实例
    /// * `Err(XlsxError)` - 创建过程中发生的验证错误
    pub fn from_dataframe(
        df: DataFrame,
        sheet_name: impl Into<String>,
        region_name: Option<String>,
        styles: RegionStyles,
    ) -> Result<Self, XlsxError> {
        let region_name = region_name.unwrap_or_else(|| "data".to_string());

        // 使用 SheetRegion::from_dataframe 创建区域
        let region = SheetRegion::from_dataframe(df, region_name, None, styles)?;

        // 创建工作表
        Self::new(sheet_name, vec![region])
    }

    /// 根据名称获取区域的不可变引用
    ///
    /// # 参数
    /// * `name` - 区域名称
    ///
    /// # 返回值
    /// * `Some(&SheetRegion)` - 找到的区域
    /// * `None` - 未找到
    pub fn region(&self, name: &str) -> Option<&SheetRegion> {
        self.regions.iter().find(|r| r.name == name)
    }

    /// 根据名称获取区域的可变引用
    ///
    /// # 参数
    /// * `name` - 区域名称
    ///
    /// # 返回值
    /// * `Some(&mut SheetRegion)` - 找到的区域
    /// * `None` - 未找到
    pub fn region_mut(&mut self, name: &str) -> Option<&mut SheetRegion> {
        self.regions.iter_mut().find(|r| r.name == name)
    }

    /// 添加区域到工作表
    ///
    /// # 参数
    /// * `region` - 要添加的区域
    ///
    /// # 返回值
    /// * `Ok(())` - 添加成功
    /// * `Err(XlsxError)` - 区域名称重复
    pub fn add_region(&mut self, region: SheetRegion) -> Result<(), XlsxError> {
        // 检查名称唯一性
        if self.regions.iter().any(|r| r.name == region.name) {
            return Err(XlsxError::GenericError(format!(
                "Region name '{}' already exists",
                region.name
            )));
        }

        self.regions.push(region);
        Ok(())
    }

    /// 获取所有区域的总行数
    ///
    /// 计算所有区域行数的总和。
    pub fn total_row_count(&self) -> usize {
        self.regions.iter().map(|r| r.row_count()).sum()
    }

    /// 获取所有区域的最大列数
    ///
    /// 找出所有区域中列数的最大值。
    pub fn max_col_count(&self) -> usize {
        self.regions
            .iter()
            .map(|r| r.col_count())
            .max()
            .unwrap_or(0)
    }

    /// 获取名称包含 "header" 的区域
    ///
    /// 根据区域名称过滤，返回名称中包含 "header"（不区分大小写）的区域。
    pub fn header_regions(&self) -> Vec<&SheetRegion> {
        self.regions
            .iter()
            .filter(|r| r.name.to_lowercase().contains("header"))
            .collect()
    }

    /// 获取名称包含 "footer" 的区域
    ///
    /// 根据区域名称过滤，返回名称中包含 "footer"（不区分大小写）的区域。
    pub fn footer_regions(&self) -> Vec<&SheetRegion> {
        self.regions
            .iter()
            .filter(|r| r.name.to_lowercase().contains("footer"))
            .collect()
    }

    /// 生成 BIFF8 数据（用于 .xls 写入）
    ///
    /// 此方法将工作表序列化为 BIFF8 格式的字节流。
    /// 所有区域按顺序写入，坐标会自动进行偏移处理。
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

        // 计算总行列数用于 DimensionsRecord
        let total_rows = self.total_row_count();
        let max_cols = self.max_col_count();

        // 创建 DimensionsRecord（总范围从 0,0 到最后一个行列）
        let dimensions = if total_rows > 0 && max_cols > 0 {
            DimensionsRecord::new(0, (total_rows - 1) as u32, 0, (max_cols - 1) as u16)
        } else {
            DimensionsRecord::default()
        };
        result.extend_from_slice(&dimensions.serialize());

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

        // 写入所有区域的数据，处理坐标偏移
        let mut current_row_offset: u32 = 0;
        for region in &self.regions {
            for (rel_row_idx, row) in region.data.iter().enumerate() {
                if row.iter().any(|c| c.is_some()) {
                    let abs_row = current_row_offset + rel_row_idx as u32;
                    result.extend_from_slice(
                        &RowRecord::from_row_data(abs_row as usize, row).serialize(),
                    );
                    result.extend_from_slice(&row_data_to_cell_records(
                        abs_row as usize,
                        row,
                        0,
                        sst,
                    ));
                }
            }
            current_row_offset += region.row_count() as u32;
        }

        result.extend_from_slice(&Window2Record::default().serialize());
        result.extend_from_slice(&EofRecord::default().serialize());

        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::Cell;
    use polars::prelude::NamedFrom;
    use polars::prelude::Series;

    #[test]
    fn test_worksheet_new() {
        let region = SheetRegion::new("data", vec![vec![Some(Cell::Text("Test".to_string()))]]);
        let sheet = WorkSheet::new("TestSheet", vec![region]).unwrap();

        assert_eq!(sheet.name, "TestSheet");
        assert_eq!(sheet.regions.len(), 1);
    }

    #[test]
    fn test_worksheet_name_validation() {
        // 空名称
        let region = SheetRegion::new("data", vec![vec![Some(Cell::Text("Test".to_string()))]]);
        let result = WorkSheet::new("", vec![region.clone()]);
        assert!(result.is_err());

        // 超过31字符
        let long_name = "a".repeat(32);
        let result = WorkSheet::new(long_name, vec![region.clone()]);
        assert!(result.is_err());

        // 包含非法字符
        let result = WorkSheet::new("Test/Sheet", vec![region.clone()]);
        assert!(result.is_err());
    }

    #[test]
    fn test_worksheet_duplicate_region_name() {
        let region1 = SheetRegion::new("data", vec![vec![Some(Cell::Text("Test1".to_string()))]]);
        let region2 = SheetRegion::new("data", vec![vec![Some(Cell::Text("Test2".to_string()))]]);
        let result = WorkSheet::new("TestSheet", vec![region1, region2]);
        assert!(result.is_err());
    }

    #[test]
    fn test_worksheet_from_dataframe() {
        let columns: Vec<polars::frame::column::Column> = vec![
            Series::new("Name".into(), vec!["Alice", "Bob"]).into(),
            Series::new("Age".into(), vec![30_i64, 25]).into(),
        ];
        let df = DataFrame::new(2, columns).unwrap();

        let sheet = WorkSheet::from_dataframe(df, "Test", None, RegionStyles::new()).unwrap();

        assert_eq!(sheet.name, "Test");
        assert_eq!(sheet.regions.len(), 1);
        assert_eq!(sheet.regions[0].name, "data");
        assert_eq!(sheet.total_row_count(), 3); // 1 header + 2 data rows
    }

    #[test]
    fn test_worksheet_region_access() {
        let region1 =
            SheetRegion::new("header", vec![vec![Some(Cell::Text("Header".to_string()))]]);
        let region2 = SheetRegion::new("data", vec![vec![Some(Cell::Text("Data".to_string()))]]);
        let sheet = WorkSheet::new("TestSheet", vec![region1, region2]).unwrap();

        // 测试 region()
        assert!(sheet.region("header").is_some());
        assert!(sheet.region("data").is_some());
        assert!(sheet.region("nonexistent").is_none());

        // 测试 region_mut()
        let mut sheet_mut = sheet;
        assert!(sheet_mut.region_mut("header").is_some());
    }

    #[test]
    fn test_worksheet_add_region() {
        let region1 = SheetRegion::new("data1", vec![vec![Some(Cell::Text("Test1".to_string()))]]);
        let mut sheet = WorkSheet::new("TestSheet", vec![region1]).unwrap();

        let region2 = SheetRegion::new("data2", vec![vec![Some(Cell::Text("Test2".to_string()))]]);
        assert!(sheet.add_region(region2).is_ok());
        assert_eq!(sheet.regions.len(), 2);

        // 添加重复名称应该失败
        let region3 = SheetRegion::new("data1", vec![vec![Some(Cell::Text("Test3".to_string()))]]);
        assert!(sheet.add_region(region3).is_err());
    }

    #[test]
    fn test_worksheet_total_row_count() {
        let region1 = SheetRegion::new(
            "r1",
            vec![
                vec![Some(Cell::Text("H1".to_string()))],
                vec![Some(Cell::Text("D1".to_string()))],
            ],
        );
        let region2 = SheetRegion::new(
            "r2",
            vec![
                vec![Some(Cell::Text("H2".to_string()))],
                vec![Some(Cell::Text("D2".to_string()))],
                vec![Some(Cell::Text("D3".to_string()))],
            ],
        );
        let sheet = WorkSheet::new("TestSheet", vec![region1, region2]).unwrap();

        assert_eq!(sheet.total_row_count(), 5);
    }

    #[test]
    fn test_worksheet_max_col_count() {
        let region1 = SheetRegion::new(
            "r1",
            vec![vec![
                Some(Cell::Text("A".to_string())),
                Some(Cell::Text("B".to_string())),
            ]],
        );
        let region2 = SheetRegion::new(
            "r2",
            vec![vec![
                Some(Cell::Text("A".to_string())),
                Some(Cell::Text("B".to_string())),
                Some(Cell::Text("C".to_string())),
            ]],
        );
        let sheet = WorkSheet::new("TestSheet", vec![region1, region2]).unwrap();

        assert_eq!(sheet.max_col_count(), 3);
    }

    #[test]
    fn test_worksheet_header_footer_regions() {
        let header_region = SheetRegion::new(
            "my_header",
            vec![vec![Some(Cell::Text("Header".to_string()))]],
        );
        let data_region =
            SheetRegion::new("data", vec![vec![Some(Cell::Text("Data".to_string()))]]);
        let footer_region = SheetRegion::new(
            "my_footer",
            vec![vec![Some(Cell::Text("Footer".to_string()))]],
        );
        let sheet =
            WorkSheet::new("TestSheet", vec![header_region, data_region, footer_region]).unwrap();

        let headers = sheet.header_regions();
        assert_eq!(headers.len(), 1);
        assert_eq!(headers[0].name, "my_header");

        let footers = sheet.footer_regions();
        assert_eq!(footers.len(), 1);
        assert_eq!(footers[0].name, "my_footer");
    }
}
