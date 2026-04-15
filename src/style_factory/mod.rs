//! 樣式工廠模組：負責將 JSON 規則轉換為 DataFrame 的單元格樣式映射。

use polars::datatypes::DataType;
use polars::prelude::*;

// 1. 聲明內部子模組
pub mod condition;
pub mod engine;
pub mod error;
pub mod rules;

// 2. 重新導出核心類型（提高外部調用的便利性）
pub use condition::StyleCondition;
pub use engine::StyleFactory;
pub use error::StyleFactoryError;
pub use rules::{ApplyAction, StyleOverride, StyleRule};

/// 內部的輔助工具函數：解析 Excel 風格的比較字符串。
/// 放在這裡可以被 engine 模組內部共享，同時保持封裝。
pub(crate) fn parse_criteria(s: &str) -> (&str, f64) {
    let s = s.trim();
    if s.starts_with(">=") {
        (">=", s[2..].parse::<f64>().unwrap_or(0.0))
    } else if s.starts_with("<=") {
        ("<=", s[2..].parse::<f64>().unwrap_or(0.0))
    } else if s.starts_with(">") {
        (">", s[1..].parse::<f64>().unwrap_or(0.0))
    } else if s.starts_with("<") {
        ("<", s[1..].parse::<f64>().unwrap_or(0.0))
    } else if s.starts_with("=") {
        ("=", s[1..].parse::<f64>().unwrap_or(0.0))
    } else {
        ("=", s.parse::<f64>().unwrap_or(0.0))
    }
}

/// 将列级别条件结果合并到行级别遮罩中
fn merge_mask(mask: &mut BooleanChunked, col_mask: BooleanChunked) {
    let other = BooleanChunked::from_iter(std::iter::once(Some(false)).chain(col_mask.into_iter()));
    *mask = &*mask & &other;
}

/// 评估样式条件并生成匹配遮罩
pub fn evaluate_row_conditions(
    df: &DataFrame,
    conditions: &[StyleCondition],
) -> Result<BooleanChunked, StyleFactoryError> {
    let height = df.height();
    let mut full_mask = BooleanChunked::full("mask".into(), true, height + 1);

    for cond in conditions {
        let mask = match cond {
            StyleCondition::Index { criteria, .. } => {
                let mut mask = vec![false; height + 1];
                for &phys_idx in criteria {
                    let df_idx = if phys_idx < 0 {
                        // 负数索引：-1 对应最后一行数据（Excel 物理行号 height）
                        // 转换为 mask 索引（包括 header）
                        let adjusted_idx = height as i32 + phys_idx + 1;
                        if adjusted_idx < 0 {
                            return Err(StyleFactoryError::IndexOutOfBounds(phys_idx, height));
                        }
                        adjusted_idx as usize
                    } else {
                        // 正数索引：直接对应 Excel 物理行号（0 = header 行）
                        phys_idx as usize
                    };

                    if df_idx <= height {
                        mask[df_idx] = true;
                    } else {
                        return Err(StyleFactoryError::IndexOutOfBounds(phys_idx, height));
                    }
                }
                Ok::<BooleanChunked, StyleFactoryError>(BooleanChunked::from_slice(
                    "mask".into(),
                    &mask,
                ))
            }
            StyleCondition::ValueRange { targets, criteria } => {
                let (op, val) = parse_criteria(criteria);
                let mut mask = vec![true; height + 1];
                mask[0] = false;
                let mut mask = BooleanChunked::new("mask".into(), &mask);

                for col_name in targets {
                    let col = df
                        .column(col_name)
                        .map_err(|_| StyleFactoryError::ColumnNotFound(col_name.clone()))?;
                    let series = col.as_materialized_series();

                    if !series.dtype().is_numeric() {
                        return Err(StyleFactoryError::TypeMismatch(
                            col_name.clone(),
                            format!("{:?}", series.dtype()),
                        ));
                    }

                    let val_lit = Series::new("lit".into(), vec![val]).cast(series.dtype())?;

                    let col_mask = match op {
                        ">" => series.gt(&val_lit)?,
                        "<" => series.lt(&val_lit)?,
                        ">=" => series.gt_eq(&val_lit)?,
                        "<=" => series.lt_eq(&val_lit)?,
                        _ => series.equal(&val_lit)?,
                    };

                    merge_mask(&mut mask, col_mask);
                }
                Ok::<BooleanChunked, StyleFactoryError>(mask)
            }
            StyleCondition::Match { targets, criteria } => {
                let mut mask = vec![true; height + 1];
                mask[0] = false;
                let mut mask = BooleanChunked::new("mask".into(), &mask);
                let criteria_set: std::collections::HashSet<&str> =
                    criteria.iter().map(|s| s.as_str()).collect();

                for col_name in targets {
                    let col = df
                        .column(col_name)
                        .map_err(|_| StyleFactoryError::ColumnNotFound(col_name.clone()))?;
                    let series = col.as_materialized_series();

                    let col_mask: Vec<bool> = if series.dtype() == &DataType::String {
                        series
                            .str()
                            .map_err(|e| StyleFactoryError::PolarsError(e))?
                            .into_iter()
                            .map(|opt_val| {
                                opt_val.map(|v| criteria_set.contains(v)).unwrap_or(false)
                            })
                            .collect()
                    } else {
                        series
                            .iter()
                            .map(|val| {
                                let s = format!("{}", val);
                                let s_clean = s.trim_matches('"');
                                criteria_set.contains(s_clean)
                            })
                            .collect()
                    };
                    merge_mask(&mut mask, BooleanChunked::new("mask".into(), col_mask));
                }
                Ok::<BooleanChunked, StyleFactoryError>(mask)
            }
            StyleCondition::Find { targets, criteria } => {
                let mut mask = vec![true; height + 1];
                mask[0] = false;
                let mut mask = BooleanChunked::new("mask".into(), &mask);

                for col_name in targets {
                    let col = df
                        .column(col_name)
                        .map_err(|_| StyleFactoryError::ColumnNotFound(col_name.clone()))?;
                    let series = col.as_materialized_series();

                    if series.dtype() != &DataType::String {
                        return Err(StyleFactoryError::TypeMismatch(
                            col_name.clone(),
                            format!("{:?}", series.dtype()),
                        ));
                    }

                    let col_mask = series
                        .str()
                        .map_err(|e| StyleFactoryError::PolarsError(e))?
                        .contains(criteria, false)
                        .map_err(|e| StyleFactoryError::PolarsError(e))?;

                    merge_mask(&mut mask, col_mask);
                }
                Ok::<BooleanChunked, StyleFactoryError>(mask)
            }
            StyleCondition::Equal { targets, criteria } => {
                let mut mask = vec![true; height + 1];
                mask[0] = false;
                let mut mask = BooleanChunked::new("mask".into(), &mask);
                if targets.len() >= 2 {
                    let first_series = df
                        .column(&targets[0])
                        .map_err(|_| StyleFactoryError::ColumnNotFound(targets[0].clone()))?
                        .as_materialized_series();
                    for col_name in &targets[1..] {
                        let other_series = df
                            .column(col_name)
                            .map_err(|_| StyleFactoryError::ColumnNotFound(col_name.clone()))?
                            .as_materialized_series();
                        let col_mask = first_series
                            .equal(other_series)
                            .map_err(|e| StyleFactoryError::PolarsError(e))?;
                        merge_mask(&mut mask, col_mask);
                    }
                }
                Ok::<BooleanChunked, StyleFactoryError>(if *criteria { mask } else { !mask })
            }
            StyleCondition::ExcludeRows { criteria, .. } => {
                let mut mask = vec![true; height + 1];
                let [p_start, p_end] = *criteria;

                let excel_start = if p_start < 0 {
                    height as i32 + p_start + 1
                } else {
                    p_start
                };

                let excel_end = if p_end < 0 {
                    height as i32 + p_end + 1
                } else {
                    p_end
                };

                if excel_start > excel_end {
                    return Ok(BooleanChunked::from_slice("mask".into(), &mask));
                }

                let actual_start = excel_start.max(0).min(height as i32) as usize;
                let actual_end = excel_end.max(0).min(height as i32) as usize;

                for i in actual_start..=actual_end {
                    if i <= height {
                        mask[i] = false;
                    }
                }

                Ok(BooleanChunked::from_slice("mask".into(), &mask))
            }
            StyleCondition::All => {
                // 返回全为 true 的掩码（包括 header 行）
                Ok(BooleanChunked::full("mask".into(), true, height + 1))
            }
        }?;
        full_mask = full_mask & mask;
    }

    Ok(full_mask)
}

/// 评估列条件并生成匹配的列索引列表
pub fn evaluate_col_conditions(
    df: &DataFrame,
    conditions: &[StyleCondition],
) -> Result<Vec<u16>, StyleFactoryError> {
    let has_data_op = conditions.iter().any(|c| {
        matches!(
            c,
            StyleCondition::ValueRange { .. } | StyleCondition::Equal { .. }
        )
    });
    if has_data_op {
        return Err(StyleFactoryError::TypeMismatch(
            "列定位".into(),
            "數據算子".into(),
        ));
    }

    let column_names = df.get_column_names();
    let mut matched = std::collections::HashSet::new();

    for cond in conditions {
        match cond {
            StyleCondition::All => {
                // 返回所有列的索引
                for idx in 0..column_names.len() {
                    matched.insert(idx as u16);
                }
            }
            c if c.get_targets().is_some() => {
                if let Some(targets) = c.get_targets() {
                    for name in targets {
                        if let Some(idx) = df.get_column_index(name) {
                            matched.insert(idx as u16);
                        }
                    }
                }
            }
            StyleCondition::Index { criteria, .. } => {
                for &idx in criteria {
                    if idx >= 0 && (idx as usize) < column_names.len() {
                        matched.insert(idx as u16);
                    }
                }
            }
            _ => {}
        }
    }
    Ok(matched.into_iter().collect())
}
