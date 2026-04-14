//! Dimension Rule 定义

use crate::style_factory::StyleCondition;
use serde::Deserialize;

/// 维度目标类型
#[derive(Deserialize, Debug, Clone, Copy, PartialEq)]
#[serde(rename_all = "snake_case")]
pub enum DimensionTarget {
    /// 行
    Row,
    /// 列
    Column,
}

/// 维度值类型
#[derive(Deserialize, Debug, Clone, Copy, PartialEq)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum DimensionValue {
    /// 固定值
    Fixed { value: f64 },
    /// 自动计算
    /// 
    /// 对于列宽：根据内容自动计算（中文字符占2宽度，英文字符占1宽度）
    /// 对于行高：根据内容自动（TODO: 待实现）
    Auto,
}

impl DimensionValue {
    /// 获取固定值，如果是 Auto 则返回 None
    pub fn fixed_value(&self) -> Option<f64> {
        match self {
            DimensionValue::Fixed { value } => Some(*value),
            DimensionValue::Auto => None,
        }
    }
    
    /// 检查是否为 Auto
    pub fn is_auto(&self) -> bool {
        matches!(self, DimensionValue::Auto)
    }
}

/// 维度规则
///
/// 定义行高或列宽的设置规则
#[derive(Deserialize, Debug, Clone)]
pub struct DimensionRule {
    /// 目标维度类型（行或列）
    pub target: DimensionTarget,
    /// 条件（复用 StyleCondition）
    pub condition: StyleCondition,
    /// 维度值
    pub value: DimensionValue,
}
