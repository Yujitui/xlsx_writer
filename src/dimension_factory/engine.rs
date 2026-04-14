//! Dimension Factory 核心实现

use crate::dimension_factory::{DimensionRule, DimensionTarget, DimensionValue};
use crate::style_factory::{evaluate_col_conditions, evaluate_row_conditions, StyleCondition};
use polars::prelude::*;
use std::collections::HashMap;

/// 维度计算结果
#[derive(Debug, Default)]
pub struct DimensionResult {
    /// 行高映射：(行索引, 高度)
    pub row_heights: HashMap<u32, f64>,
    /// 列宽映射：(列索引, 宽度)
    pub col_widths: HashMap<u16, f64>,
}

/// 维度工厂
///
/// 负责根据规则计算行高和列宽
pub struct DimensionFactory {
    /// 规则集
    pub rules: Vec<DimensionRule>,
}

impl DimensionFactory {
    /// 创建新的维度工厂实例
    pub fn new(value: serde_json::Value) -> Result<Self, Box<dyn std::error::Error>> {
        let rules: Vec<DimensionRule> = if value.is_array() {
            serde_json::from_value(value)?
        } else {
            if !value.is_null() {
                eprintln!(
                    "Warning: DimensionFactory input is not a valid array, using empty rules."
                );
            }
            vec![]
        };

        Ok(Self { rules })
    }

    /// 从 JSON 字符串创建维度工厂实例
    pub fn from_json_str(json_str: &str) -> Result<Self, Box<dyn std::error::Error>> {
        let value: serde_json::Value = serde_json::from_str(json_str)?;
        Self::new(value)
    }

    /// 执行维度规则评估并生成行高列宽映射表
    pub fn execute(&self, df: &DataFrame) -> Result<DimensionResult, Box<dyn std::error::Error>> {
        let mut result = DimensionResult::default();

        for rule in &self.rules {
            match rule.target {
                DimensionTarget::Row => {
                    self.apply_row_rule(df, rule, &mut result)?;
                }
                DimensionTarget::Column => {
                    self.apply_column_rule(df, rule, &mut result)?;
                }
            }
        }

        Ok(result)
    }

    /// 应用行高规则
    fn apply_row_rule(
        &self,
        df: &DataFrame,
        rule: &DimensionRule,
        result: &mut DimensionResult,
    ) -> Result<(), Box<dyn std::error::Error>> {
        // 使用公共函数评估行条件
        let row_mask = evaluate_row_conditions(df, &[rule.condition.clone()])
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

        // 计算行高值
        let height = match rule.value {
            DimensionValue::Fixed { value } => value,
            DimensionValue::Auto => {
                // TODO: 根据内容自动计算行高
                // 暂时使用默认行高
                15.0
            }
        };

        // 应用行高到匹配的行
        // row_mask 长度为 height + 1，其中索引 0 是 header 行，1..=height 是数据行
        for (r_idx, matched) in row_mask.into_iter().enumerate() {
            if let Some(true) = matched {
                // 跳过 header 行（索引 0）
                if r_idx > 0 {
                    // 数据行索引减 1 转换为 0-based
                    result.row_heights.insert((r_idx - 1) as u32, height);
                }
            }
        }

        Ok(())
    }

    /// 应用列宽规则
    fn apply_column_rule(
        &self,
        df: &DataFrame,
        rule: &DimensionRule,
        result: &mut DimensionResult,
    ) -> Result<(), Box<dyn std::error::Error>> {
        // 检查列条件是否包含不支持的数据驱动型条件
        if let StyleCondition::ValueRange { .. } | StyleCondition::Equal { .. } = rule.condition {
            // 对于列宽规则，这些条件类型被静默忽略
            return Ok(());
        }

        // 使用公共函数评估列条件
        let col_indices = evaluate_col_conditions(df, &[rule.condition.clone()])
            .map_err(|e| Box::new(e) as Box<dyn std::error::Error>)?;

        // 计算列宽值
        match rule.value {
            DimensionValue::Fixed { value } => {
                // 固定值：所有列使用相同宽度
                for col_idx in col_indices {
                    result.col_widths.insert(col_idx, value);
                }
            }
            DimensionValue::Auto => {
                // 自动计算：每列根据自己的内容计算宽度
                for col_idx in &col_indices {
                    let width = self.calculate_single_col_width(df, *col_idx);
                    result.col_widths.insert(*col_idx, width);
                }
            }
        };

        Ok(())
    }

    /// 计算单列的自动宽度
    ///
    /// 遍历该列的所有单元格，计算内容宽度：
    /// - 中文字符（CJK Unified Ideographs 范围）计 2 宽度单位
    /// - 其他字符计 1 宽度单位
    /// - 返回最大值（至少为 8.0）
    fn calculate_single_col_width(&self, df: &DataFrame, col_idx: u16) -> f64 {
        let col_idx = col_idx as usize;
        let columns: Vec<_> = df.columns().into_iter().collect();

        if col_idx >= columns.len() {
            return 8.0;
        }

        let col = &columns[col_idx];
        let mut max_content_width: f64 = 0.0;

        // 遍历该列的所有值
        for i in 0..col.len() {
            if let Ok(val) = col.get(i) {
                // 提取字符串内容
                let s = match &val {
                    AnyValue::String(s) => s.to_string(),
                    _ => format!("{}", val),
                };
                let width = self.calculate_string_width(&s);
                max_content_width = max_content_width.max(width as f64);
            }
        }

        // 确保最小宽度，并添加内边距
        (max_content_width + 2.0).max(8.0)
    }

    /// 计算字符串宽度
    ///
    /// 中文字符（Unicode 范围 0x4E00-0x9FFF）计 2 宽度单位
    /// 其他字符计 1 宽度单位
    fn calculate_string_width(&self, s: &str) -> usize {
        s.chars()
            .map(|c| {
                if (0x4E00..=0x9FFF).contains(&(c as u32)) {
                    2 // 中文字符
                } else {
                    1 // 其他字符
                }
            })
            .sum()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_calculate_string_width() {
        let factory = DimensionFactory { rules: vec![] };

        // Test Chinese
        let chinese_width = factory.calculate_string_width("张三");
        assert_eq!(chinese_width, 4);

        // Test English
        let english_width = factory.calculate_string_width("AB");
        assert_eq!(english_width, 2);
    }
}

#[test]
fn test_debug_auto_width() {
    use polars::prelude::*;

    let df = df! {
        "name" => ["张三", "李四"],
        "code" => ["AB", "CD"]
    }
    .expect("Failed");

    let factory = DimensionFactory { rules: vec![] };

    println!(
        "Column 0 width: {}",
        factory.calculate_single_col_width(&df, 0)
    );
    println!(
        "Column 1 width: {}",
        factory.calculate_single_col_width(&df, 1)
    );

    // Check what's in the columns
    for (idx, col) in df.columns().into_iter().enumerate() {
        println!("Column {} ({}):", idx, col.name());
        for i in 0..col.len() {
            if let Ok(val) = col.get(i) {
                match &val {
                    AnyValue::String(s) => {
                        println!(
                            "  Row {}: String({}) -> width {}",
                            i,
                            s,
                            factory.calculate_string_width(s)
                        );
                        s.to_string()
                    }
                    _ => {
                        let s = format!("{}", val);
                        println!(
                            "  Row {}: Other({}) -> width {}",
                            i,
                            s,
                            factory.calculate_string_width(&s)
                        );
                        s
                    }
                };
            }
        }
    }
}
