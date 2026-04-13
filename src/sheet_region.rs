//! 工作表特殊区域模块
//!
//! 本模块定义了 Excel 工作表中的特殊区域结构，用于存储表头、脚注等
//! 独立于主数据区域的内容。

use crate::cell::Cell;
use std::collections::HashMap;
use std::sync::Arc;

/// 特殊区域类型
#[derive(Debug, Clone, PartialEq)]
pub enum RegionType {
    /// 表头（显示在主数据上方）
    Header,
    /// 脚注（显示在主数据下方）
    Footer,
}

/// 工作表特殊区域
///
/// 用于存储表头、脚注等特殊区域的数据和样式。
/// 区域使用相对坐标（以区域自身为起点），写入时自动计算实际位置。
#[derive(Debug, Clone)]
pub struct SheetRegion {
    /// 区域类型
    pub region_type: RegionType,
    /// 区域数据：二维单元格
    /// 外层 Vec 是行，内层 Vec 是列
    pub data: Vec<Vec<Option<Cell>>>,
    /// 样式映射（相对坐标，以区域左上角为原点）
    /// Key: (row, col) - 行列坐标 (0-based)
    /// Value: 样式名称
    pub style_map: Option<HashMap<(u32, u16), Arc<str>>>,
    /// 合并区域（相对坐标）
    /// 每个元组表示一个合并区域：(起始行, 起始列, 结束行, 结束列)
    pub merge_ranges: Option<Vec<(u32, u16, u32, u16)>>,
}

impl SheetRegion {
    /// 创建新的特殊区域
    ///
    /// # 参数
    /// * `region_type` - 区域类型（Header 或 Footer）
    /// * `data` - 区域数据
    ///
    /// # 返回值
    /// 返回一个新的 SheetRegion 实例
    pub fn new(region_type: RegionType, data: Vec<Vec<Option<Cell>>>) -> Self {
        SheetRegion {
            region_type,
            data,
            style_map: None,
            merge_ranges: None,
        }
    }

    /// 设置样式映射
    ///
    /// # 参数
    /// * `style_map` - 样式映射表
    pub fn with_style_map(mut self, style_map: HashMap<(u32, u16), Arc<str>>) -> Self {
        self.style_map = Some(style_map);
        self
    }

    /// 设置合并区域
    ///
    /// # 参数
    /// * `merge_ranges` - 合并区域列表
    pub fn with_merge_ranges(mut self, merge_ranges: Vec<(u32, u16, u32, u16)>) -> Self {
        self.merge_ranges = Some(merge_ranges);
        self
    }

    /// 获取区域的行数
    pub fn row_count(&self) -> usize {
        self.data.len()
    }

    /// 获取区域的最大列数
    pub fn col_count(&self) -> usize {
        self.data.iter().map(|row| row.len()).max().unwrap_or(0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sheet_region_creation() {
        let data = vec![
            vec![Some(Cell::Text("Header1".to_string()))],
            vec![Some(Cell::Text("Header2".to_string()))],
        ];
        let region = SheetRegion::new(RegionType::Header, data);

        assert_eq!(region.region_type, RegionType::Header);
        assert_eq!(region.row_count(), 2);
        assert_eq!(region.col_count(), 1);
    }

    #[test]
    fn test_sheet_region_with_styles() {
        let data = vec![vec![Some(Cell::Text("Test".to_string()))]];
        let mut style_map = HashMap::new();
        style_map.insert((0, 0), Arc::from("header_style"));

        let region = SheetRegion::new(RegionType::Footer, data).with_style_map(style_map);

        assert!(region.style_map.is_some());
        assert_eq!(
            region.style_map.as_ref().unwrap().get(&(0, 0)),
            Some(&Arc::from("header_style"))
        );
    }
}
