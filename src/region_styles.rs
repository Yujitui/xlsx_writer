//! Region 级别的样式属性模块
//!
//! 本模块定义了 SheetRegion 使用的完整样式属性结构，
//! 包含单元格样式、行高、列宽等所有视觉属性。

use std::collections::HashMap;
use std::sync::Arc;

/// Region 级别的样式属性
///
/// 包含单元格样式、行列尺寸、合并区域等所有视觉属性。
/// 这是一个扁平化结构，便于序列化和操作。
#[derive(Debug, Clone, Default)]
pub struct RegionStyles {
    /// 单元格样式映射 (row, col) -> style_name
    pub cell_styles: HashMap<(u32, u16), Arc<str>>,

    /// 特定行高（row -> height in points）
    /// 预留，暂不启用
    pub row_heights: HashMap<u32, f64>,

    /// 特定列宽（col -> width in characters）
    /// 预留，暂不启用
    /// 存在的列使用指定宽度，不存在的列自动计算
    pub col_widths: HashMap<u16, f64>,

    /// 合并区域列表 (start_row, start_col, end_row, end_col)
    /// 坐标基于 0-based 系统
    pub merge_ranges: Vec<(u32, u16, u32, u16)>,
}

impl RegionStyles {
    /// 创建空的 RegionStyles
    pub fn new() -> Self {
        Self::default()
    }

    /// 坐标调整（用于 include_header=false）
    ///
    /// 当不导出表头时，所有行坐标需要减1：
    /// - row=0 的样式被删除（原表头行）
    /// - row>0 的样式行号减1
    /// - 涉及 row=0 的合并区域被删除
    /// - 其他合并区域的行号减1
    pub fn adjust_coordinates(&mut self) {
        // 调整 cell_styles: row=0 删除，其他 row-1
        self.cell_styles = self
            .cell_styles
            .drain()
            .filter(|((row, _), _)| *row != 0)
            .map(|((row, col), style)| ((row - 1, col), style))
            .collect();

        // 调整 row_heights: row=0 删除，其他 row-1
        self.row_heights = self
            .row_heights
            .drain()
            .filter(|(row, _)| *row != 0)
            .map(|(row, height)| (row - 1, height))
            .collect();

        // 调整 merge_ranges: 涉及 row=0 的删除，其他 row-1
        self.merge_ranges = self
            .merge_ranges
            .drain(..)
            .filter(|(start_row, _, end_row, _)| *start_row != 0 && *end_row != 0)
            .map(|(start_row, start_col, end_row, end_col)| {
                (start_row - 1, start_col, end_row - 1, end_col)
            })
            .collect();

        // col_widths 列坐标不受影响（列与表头无关）
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_region_styles_new() {
        let styles = RegionStyles::new();
        assert!(styles.cell_styles.is_empty());
        assert!(styles.row_heights.is_empty());
        assert!(styles.col_widths.is_empty());
        assert!(styles.merge_ranges.is_empty());
    }

    #[test]
    fn test_adjust_coordinates() {
        let mut styles = RegionStyles::new();
        styles.cell_styles.insert((0, 0), Arc::from("header_style"));
        styles.cell_styles.insert((1, 0), Arc::from("row1_style"));
        styles.cell_styles.insert((2, 1), Arc::from("row2_style"));
        styles.adjust_coordinates();

        // row=0 (header) 被删除，row=1 变为 row=0
        assert_eq!(
            styles.cell_styles.get(&(0, 0)),
            Some(&Arc::from("row1_style"))
        );

        // row=2 变为 row=1
        assert_eq!(
            styles.cell_styles.get(&(1, 1)),
            Some(&Arc::from("row2_style"))
        );
    }

    #[test]
    fn test_adjust_coordinates_with_row_heights() {
        let mut styles = RegionStyles::new();
        styles.cell_styles.insert((0, 0), Arc::from("header"));
        styles.cell_styles.insert((1, 0), Arc::from("data"));
        styles.row_heights.insert(0, 30.0);
        styles.row_heights.insert(1, 25.0);
        styles.row_heights.insert(2, 20.0);

        styles.adjust_coordinates();

        // cell_styles 调整：row=0 (header) 被删除，row=1 变为 row=0
        assert_eq!(styles.cell_styles.get(&(0, 0)), Some(&Arc::from("data"))); // 原 row=1

        // row_heights 调整：row=0 (header, 30.0) 被删除，row=1 (25.0) 变为 row=0，row=2 (20.0) 变为 row=1
        assert_eq!(styles.row_heights.get(&0), Some(&25.0)); // 原 row=1 变为 row=0
        assert_eq!(styles.row_heights.get(&1), Some(&20.0)); // 原 row=2 变为 row=1
        assert_eq!(styles.row_heights.len(), 2); // 原 row=0 被删除
    }

    #[test]
    fn test_adjust_coordinates_with_merge_ranges() {
        let mut styles = RegionStyles::new();
        // 添加一些样式
        styles.cell_styles.insert((0, 0), Arc::from("header"));
        styles.cell_styles.insert((1, 0), Arc::from("data1"));
        styles.cell_styles.insert((2, 0), Arc::from("data2"));

        // 添加合并区域
        styles.merge_ranges.push((0, 0, 0, 3)); // 表头合并，应该被删除
        styles.merge_ranges.push((1, 0, 3, 0)); // 华东合并，应该调整到 (0, 0, 2, 0)
        styles.merge_ranges.push((4, 0, 5, 0)); // 华北合并，应该调整到 (3, 0, 4, 0)

        styles.adjust_coordinates();

        // cell_styles 调整
        assert_eq!(styles.cell_styles.get(&(0, 0)), Some(&Arc::from("data1"))); // 原 row=1
        assert_eq!(styles.cell_styles.get(&(1, 0)), Some(&Arc::from("data2"))); // 原 row=2

        // merge_ranges 调整
        assert_eq!(styles.merge_ranges.len(), 2); // 原 row=0 的合并被删除
        assert!(styles.merge_ranges.contains(&(0, 0, 2, 0))); // 原 (1, 0, 3, 0) 变为 (0, 0, 2, 0)
        assert!(styles.merge_ranges.contains(&(3, 0, 4, 0))); // 原 (4, 0, 5, 0) 变为 (3, 0, 4, 0)
    }
}
