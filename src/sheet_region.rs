//! 工作表区域模块
//!
//! 本模块定义了 Excel 工作表中的区域结构。
//! 每个区域是自包含的数据单元，包含数据、样式和合并区域信息。

use crate::cell::Cell;
use crate::error::XlsxError;
use polars::prelude::{DataFrame, Series};
use std::collections::HashMap;
use std::sync::Arc;

/// 工作表区域
#[derive(Debug, Clone)]
pub struct SheetRegion {
    pub name: String,
    pub data: Vec<Vec<Option<Cell>>>,
    style_map: Option<HashMap<(u32, u16), Arc<str>>>,
    merge_ranges: Option<Vec<(u32, u16, u32, u16)>>,
}

impl SheetRegion {
    pub fn new(name: impl Into<String>, data: Vec<Vec<Option<Cell>>>) -> Self {
        SheetRegion {
            name: name.into(),
            data,
            style_map: None,
            merge_ranges: None,
        }
    }

    pub fn empty(name: impl Into<String>) -> Self {
        SheetRegion {
            name: name.into(),
            data: Vec::new(),
            style_map: None,
            merge_ranges: None,
        }
    }

    pub fn with_style_map(mut self, style_map: HashMap<(u32, u16), Arc<str>>) -> Self {
        self.style_map = Some(style_map);
        self
    }

    pub fn with_merge_ranges(mut self, merge_ranges: Vec<(u32, u16, u32, u16)>) -> Self {
        self.merge_ranges = Some(merge_ranges);
        self
    }

    pub fn row_count(&self) -> usize {
        self.data.len()
    }

    pub fn data_row_count(&self) -> usize {
        self.data.len().saturating_sub(1)
    }

    pub fn col_count(&self) -> usize {
        self.data.iter().map(|row| row.len()).max().unwrap_or(0)
    }

    pub fn column_names(&self) -> Vec<String> {
        self.data
            .first()
            .map(|header_row| {
                header_row
                    .iter()
                    .enumerate()
                    .map(|(idx, cell)| match cell {
                        Some(Cell::Text(s)) if !s.is_empty() => s.clone(),
                        _ => format!("Column_{}", idx),
                    })
                    .collect()
            })
            .unwrap_or_default()
    }

    pub fn to_dataframe(&self) -> Result<DataFrame, XlsxError> {
        use polars::datatypes::AnyValue;

        if self.data.len() <= 1 {
            return Err(XlsxError::EmptyDataFrame);
        }

        let column_names = self.column_names();
        let mut series_vec: Vec<Series> = Vec::new();

        for (col_idx, col_name) in column_names.iter().enumerate() {
            let mut data: Vec<AnyValue> = Vec::new();

            for row_idx in 1..self.data.len() {
                let cell = self.data[row_idx].get(col_idx).and_then(|opt| opt.as_ref());

                let any_val = match cell {
                    Some(Cell::Text(s)) => AnyValue::StringOwned(s.clone().into()),
                    Some(Cell::Number(n)) => AnyValue::Float64(*n),
                    Some(Cell::Boolean(b)) => AnyValue::Boolean(*b),
                    None => AnyValue::Null,
                };
                data.push(any_val);
            }

            let series = Series::from_any_values(col_name.as_str().into(), &data, true)
                .map_err(|e| XlsxError::GenericError(format!("Series error: {}", e)))?;
            series_vec.push(series);
        }

        let columns: Vec<polars::prelude::Column> =
            series_vec.into_iter().map(|s| s.into()).collect();

        let height = if columns.is_empty() {
            0
        } else {
            columns[0].len()
        };
        DataFrame::new(height, columns)
            .map_err(|e| XlsxError::GenericError(format!("DataFrame error: {}", e)))
    }

    pub fn from_dataframe(
        df: DataFrame,
        name: impl Into<String>,
        include_header: Option<bool>,
        style_map: Option<HashMap<(u32, u16), Arc<str>>>,
        merge_ranges: Option<Vec<(u32, u16, u32, u16)>>,
    ) -> Result<Self, XlsxError> {
        if df.height() == 0 || df.width() == 0 {
            return Err(XlsxError::EmptyDataFrame);
        }

        let include_header = include_header.unwrap_or(true);

        let column_names: Vec<String> = df
            .get_column_names()
            .iter()
            .map(|s| s.to_string())
            .collect();

        let mut cells: Vec<Vec<Option<Cell>>> = Vec::new();

        if include_header {
            cells.push(
                column_names
                    .iter()
                    .map(|name| Some(Cell::Text(name.clone())))
                    .collect(),
            );
        }

        for row_idx in 0..df.height() {
            let mut row: Vec<Option<Cell>> = Vec::new();
            for col in df.columns() {
                let val: polars::datatypes::AnyValue = col.get(row_idx).map_err(|e| {
                    XlsxError::GenericError(format!(
                        "Get value error at ({}, {:?}): {}",
                        row_idx,
                        col.name(),
                        e
                    ))
                })?;
                row.push(Cell::from_any_value(&val));
            }
            cells.push(row);
        }

        let (final_style_map, final_merge_ranges) = if include_header {
            (style_map, merge_ranges)
        } else {
            // 当 include_header=false 时，所有坐标需要减1（因为删除了表头行）
            // 但 row=0 的样式现在是给第0行数据的，不应该被过滤
            let adjusted_style_map = style_map.map(|map| {
                map.into_iter()
                    .filter(|((row, _), _)| *row != 0) // 删除原来表头(row=0)的样式
                    .map(|((row, col), style)| ((row - 1, col), style)) // 数据行坐标减1
                    .collect()
            });

            let adjusted_merge_ranges = merge_ranges.map(|ranges| {
                ranges
                    .into_iter()
                    .filter(|(start_row, _, _, _)| *start_row != 0) // 删除涉及表头的合并
                    .map(|(start_row, start_col, end_row, end_col)| {
                        (start_row - 1, start_col, end_row - 1, end_col) // 坐标减1
                    })
                    .collect()
            });

            (adjusted_style_map, adjusted_merge_ranges)
        };

        Ok(SheetRegion {
            name: name.into(),
            data: cells,
            style_map: final_style_map,
            merge_ranges: final_merge_ranges,
        })
    }

    pub fn set_style(&mut self, row: u32, col: u16, style: impl Into<Arc<str>>) {
        let row_usize = row as usize;
        let col_usize = col as usize;
        if row_usize >= self.data.len() || col_usize >= self.data[row_usize].len() {
            return;
        }
        let style_map = self.style_map.get_or_insert_with(HashMap::new);
        style_map.insert((row, col), style.into());
    }

    pub fn clear_style(&mut self, row: u32, col: u16) {
        if let Some(style_map) = &mut self.style_map {
            style_map.remove(&(row, col));
            if style_map.is_empty() {
                self.style_map = None;
            }
        }
    }

    pub fn clear_all_styles(&mut self) {
        self.style_map = None;
    }

    pub fn get_style(&self, row: u32, col: u16) -> Option<&Arc<str>> {
        self.style_map.as_ref().and_then(|map| map.get(&(row, col)))
    }

    pub fn styles(&self) -> Option<&HashMap<(u32, u16), Arc<str>>> {
        self.style_map.as_ref()
    }

    pub fn add_merge(&mut self, start_row: u32, start_col: u16, end_row: u32, end_col: u16) {
        if start_row > end_row || start_col > end_col {
            return;
        }
        let start_row_usize = start_row as usize;
        let end_row_usize = end_row as usize;
        let start_col_usize = start_col as usize;
        let _end_col_usize = end_col as usize;
        if start_row_usize >= self.data.len() {
            return;
        }
        let max_cols = self.data[start_row_usize].len();
        if start_col_usize >= max_cols {
            return;
        }
        if end_row_usize >= self.data.len() {
            return;
        }
        let merges = self.merge_ranges.get_or_insert_with(Vec::new);
        merges.push((start_row, start_col, end_row, end_col));
    }

    pub fn clear_merge_at(&mut self, row: u32, col: u16) {
        if let Some(merges) = &mut self.merge_ranges {
            merges.retain(|(start_row, start_col, end_row, end_col)| {
                !(row >= *start_row && row <= *end_row && col >= *start_col && col <= *end_col)
            });
            if merges.is_empty() {
                self.merge_ranges = None;
            }
        }
    }

    pub fn clear_all_merges(&mut self) {
        self.merge_ranges = None;
    }

    pub fn is_merged(&self, row: u32, col: u16) -> bool {
        self.get_merge(row, col).is_some()
    }

    pub fn get_merge(&self, row: u32, col: u16) -> Option<(u32, u16, u32, u16)> {
        self.merge_ranges.as_ref().and_then(|merges| {
            merges
                .iter()
                .find_map(|(start_row, start_col, end_row, end_col)| {
                    if row >= *start_row && row <= *end_row && col >= *start_col && col <= *end_col
                    {
                        Some((*start_row, *start_col, *end_row, *end_col))
                    } else {
                        None
                    }
                })
        })
    }

    pub fn merges(&self) -> Option<&Vec<(u32, u16, u32, u16)>> {
        self.merge_ranges.as_ref()
    }

    pub fn visualize(&self) -> String {
        let row_count = self.data.len();
        let col_count = self.col_count();

        if row_count == 0 || col_count == 0 {
            return format!("SheetRegion '{}' (empty)", self.name);
        }

        let max_display_rows = 20;
        let max_display_cols = 10;

        let mut lines = Vec::new();
        lines.push(format!(
            "SheetRegion '{}' ({} rows x {} cols)",
            self.name, row_count, col_count
        ));

        let rows_to_display: Vec<usize> = if row_count <= max_display_rows {
            (0..row_count).collect()
        } else {
            let mut rows: Vec<usize> = (0..10).collect();
            rows.push(usize::MAX);
            rows.extend((row_count - 10)..row_count);
            rows
        };

        for row_idx in &rows_to_display {
            if *row_idx == usize::MAX {
                lines.push("  ...".to_string());
                continue;
            }

            let row = &self.data[*row_idx];
            let mut cells = Vec::new();

            let cols_to_display: Vec<usize> = if col_count <= max_display_cols {
                (0..col_count).collect()
            } else {
                let mut cols: Vec<usize> = (0..5).collect();
                cols.push(usize::MAX);
                cols.extend((col_count - 5)..col_count);
                cols
            };

            for col_idx in &cols_to_display {
                if *col_idx == usize::MAX {
                    cells.push("...".to_string());
                    continue;
                }

                let cell_str = row
                    .get(*col_idx)
                    .and_then(|opt| opt.as_ref())
                    .map(|c| c.to_string())
                    .unwrap_or_else(|| "".to_string());

                let truncated = if cell_str.len() > 15 {
                    format!("{}...", &cell_str[..12])
                } else {
                    cell_str
                };
                cells.push(truncated);
            }

            lines.push(format!("  [{}] {}", row_idx, cells.join(" | ")));
        }

        lines.join("\n")
    }

    pub fn visualize_compact(&self) -> String {
        format!(
            "SheetRegion '{}' ({} rows x {} cols, {} styles, {} merges)",
            self.name,
            self.row_count(),
            self.col_count(),
            self.style_map.as_ref().map(|m| m.len()).unwrap_or(0),
            self.merge_ranges.as_ref().map(|m| m.len()).unwrap_or(0)
        )
    }

    pub fn validate(&self) -> Vec<String> {
        let mut issues = Vec::new();
        let row_count = self.data.len();

        if row_count == 0 {
            issues.push("Region has no data rows".to_string());
            return issues;
        }

        let col_count = self.col_count();
        if col_count == 0 {
            issues.push("Region has no columns".to_string());
        }

        if let Some(style_map) = &self.style_map {
            for ((row, col), _) in style_map.iter() {
                let row_usize = *row as usize;
                let col_usize = *col as usize;
                if row_usize >= row_count {
                    issues.push(format!("Style at out-of-bounds row: ({}, {})", row, col));
                } else if col_usize >= self.data[row_usize].len() {
                    issues.push(format!("Style at out-of-bounds col: ({}, {})", row, col));
                }
            }
        }

        if let Some(merges) = &self.merge_ranges {
            for (i, (start_row, start_col, end_row, end_col)) in merges.iter().enumerate() {
                if start_row > end_row || start_col > end_col {
                    issues.push(format!(
                        "Merge {} has invalid range: ({},{})-({},{})",
                        i, start_row, start_col, end_row, end_col
                    ));
                }
                let start_row_usize = *start_row as usize;
                let end_row_usize = *end_row as usize;
                if start_row_usize >= row_count || end_row_usize >= row_count {
                    issues.push(format!("Merge {} is out of row bounds", i));
                }
            }
        }

        issues
    }

    pub fn validate_and_fix(&mut self) -> Vec<String> {
        let mut fixes = Vec::new();
        let row_count = self.data.len();

        if row_count == 0 {
            return vec!["Region is empty, nothing to fix".to_string()];
        }

        if let Some(style_map) = &mut self.style_map {
            let before_count = style_map.len();
            style_map.retain(|(row, col), _| {
                let row_usize = *row as usize;
                let col_usize = *col as usize;
                row_usize < row_count && col_usize < self.data[row_usize].len()
            });
            let after_count = style_map.len();
            if before_count != after_count {
                fixes.push(format!(
                    "Removed {} out-of-bounds styles",
                    before_count - after_count
                ));
            }
            if style_map.is_empty() {
                self.style_map = None;
            }
        }

        if let Some(merges) = &mut self.merge_ranges {
            let before_count = merges.len();
            merges.retain(|(start_row, start_col, end_row, end_col)| {
                if start_row > end_row || start_col > end_col {
                    return false;
                }
                let start_row_usize = *start_row as usize;
                let end_row_usize = *end_row as usize;
                start_row_usize < row_count && end_row_usize < row_count
            });
            let after_count = merges.len();
            if before_count != after_count {
                fixes.push(format!(
                    "Removed {} invalid merges",
                    before_count - after_count
                ));
            }
            if merges.is_empty() {
                self.merge_ranges = None;
            }
        }

        fixes
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sheet_region_creation() {
        let data = vec![
            vec![
                Some(Cell::Text("Name".to_string())),
                Some(Cell::Text("Age".to_string())),
            ],
            vec![
                Some(Cell::Text("Alice".to_string())),
                Some(Cell::Number(30.0)),
            ],
        ];
        let region = SheetRegion::new("test_data", data);

        assert_eq!(region.name, "test_data");
        assert_eq!(region.row_count(), 2);
        assert_eq!(region.data_row_count(), 1);
        assert_eq!(region.col_count(), 2);
    }

    #[test]
    fn test_sheet_region_empty() {
        let region = SheetRegion::empty("empty_region");
        assert_eq!(region.name, "empty_region");
        assert_eq!(region.row_count(), 0);
        assert_eq!(region.col_count(), 0);
    }

    #[test]
    fn test_sheet_region_column_names() {
        let data = vec![
            vec![
                Some(Cell::Text("Name".to_string())),
                Some(Cell::Text("Age".to_string())),
            ],
            vec![
                Some(Cell::Text("Alice".to_string())),
                Some(Cell::Number(30.0)),
            ],
        ];
        let region = SheetRegion::new("test", data);
        assert_eq!(region.column_names(), vec!["Name", "Age"]);
    }

    #[test]
    fn test_sheet_region_with_styles() {
        let data = vec![vec![Some(Cell::Text("Test".to_string()))]];
        let mut style_map = HashMap::new();
        style_map.insert((0, 0), Arc::from("header_style"));

        let region = SheetRegion::new("styled", data).with_style_map(style_map);

        assert!(region.styles().is_some());
        assert_eq!(region.get_style(0, 0), Some(&Arc::from("header_style")));
    }

    #[test]
    fn test_style_management() {
        let data = vec![
            vec![
                Some(Cell::Text("A".to_string())),
                Some(Cell::Text("B".to_string())),
            ],
            vec![
                Some(Cell::Text("C".to_string())),
                Some(Cell::Text("D".to_string())),
            ],
        ];
        let mut region = SheetRegion::new("test", data);

        region.set_style(0, 0, "header");
        assert_eq!(region.get_style(0, 0), Some(&Arc::from("header")));

        region.clear_style(0, 0);
        assert_eq!(region.get_style(0, 0), None);

        region.set_style(100, 0, "out_of_bounds");
        assert_eq!(region.get_style(100, 0), None);
    }

    #[test]
    fn test_merge_management() {
        let data = vec![
            vec![
                Some(Cell::Text("A".to_string())),
                Some(Cell::Text("B".to_string())),
            ],
            vec![
                Some(Cell::Text("C".to_string())),
                Some(Cell::Text("D".to_string())),
            ],
        ];
        let mut region = SheetRegion::new("test", data);

        region.add_merge(0, 0, 1, 1);
        assert!(region.is_merged(0, 0));
        assert!(region.is_merged(1, 1));

        let merge = region.get_merge(0, 0);
        assert_eq!(merge, Some((0, 0, 1, 1)));

        region.clear_merge_at(0, 0);
        assert!(!region.is_merged(0, 0));
    }
}
