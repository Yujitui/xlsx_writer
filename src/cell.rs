//! 通用 Excel 单元格类型
//!
//! 本模块定义了 Excel 工作表中通用的单元格数据类型，
//! 可被 xlsx 和 xls 格式共用。

use polars::datatypes::AnyValue;

/// 通用 Excel 单元格类型
///
/// 封装了 Excel 单元格可以包含的不同数据类型。
/// 这是一个简化版本，专注于数据导出场景，暂不支持公式和富文本。
#[derive(Debug, Clone, PartialEq)]
pub enum Cell {
    /// 数值类型（f64）
    ///
    /// Excel 内部将数字存储为 64 位浮点数。
    /// 这包括整数、小数以及日期（以序列号形式存储）。
    Number(f64),

    /// 文本类型
    ///
    /// 纯文本内容。注意：暂不支持富文本（部分格式）。
    Text(String),

    /// 布尔类型
    ///
    /// 逻辑值 true 或 false。
    Boolean(bool),
}

impl Cell {
    /// 将单元格内容转换为字符串
    pub fn to_string(&self) -> String {
        match self {
            Cell::Number(n) => n.to_string(),
            Cell::Text(s) => s.clone(),
            Cell::Boolean(b) => b.to_string(),
        }
    }

    /// 从 Polars AnyValue 转换
    ///
    /// # 参数
    /// * `value` - Polars 的 AnyValue
    ///
    /// # 返回值
    /// * `Some(Cell)` - 成功转换
    /// * `None` - 不支持的类型或 Null
    pub fn from_any_value(value: &AnyValue) -> Option<Self> {
        match value {
            AnyValue::String(s) => Some(Cell::Text(s.to_string())),
            AnyValue::StringOwned(s) => Some(Cell::Text(s.to_string())),
            AnyValue::Float64(n) => Some(Cell::Number(*n)),
            AnyValue::Float32(n) => Some(Cell::Number(*n as f64)),
            AnyValue::Int8(v) => Some(Cell::Number(*v as f64)),
            AnyValue::Int16(v) => Some(Cell::Number(*v as f64)),
            AnyValue::Int32(v) => Some(Cell::Number(*v as f64)),
            AnyValue::Int64(v) => Some(Cell::Number(*v as f64)),
            AnyValue::UInt8(v) => Some(Cell::Number(*v as f64)),
            AnyValue::UInt16(v) => Some(Cell::Number(*v as f64)),
            AnyValue::UInt32(v) => Some(Cell::Number(*v as f64)),
            AnyValue::UInt64(v) => Some(Cell::Number(*v as f64)),
            AnyValue::Boolean(b) => Some(Cell::Boolean(*b)),
            AnyValue::Date(days) => Some(Cell::Number(*days as f64)),
            AnyValue::Datetime(v, unit, _) => {
                // 转换为 Excel 日期序列号（从 1899-12-30 起的天数）
                let days = match unit {
                    polars::datatypes::TimeUnit::Milliseconds => *v as f64 / 86_400_000.0,
                    polars::datatypes::TimeUnit::Microseconds => *v as f64 / 86_400_000_000.0,
                    polars::datatypes::TimeUnit::Nanoseconds => *v as f64 / 86_400_000_000_000.0,
                };
                Some(Cell::Number(days))
            }
            AnyValue::Null => None,
            _ => {
                // 其他类型转为字符串作为兜底
                Some(Cell::Text(format!("{}", value)))
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_to_string() {
        assert_eq!(Cell::Number(42.0).to_string(), "42");
        assert_eq!(Cell::Text("hello".to_string()).to_string(), "hello");
        assert_eq!(Cell::Boolean(true).to_string(), "true");
    }

    #[test]
    fn test_from_any_value() {
        assert_eq!(
            Cell::from_any_value(&AnyValue::Float64(3.14)),
            Some(Cell::Number(3.14))
        );
        assert_eq!(
            Cell::from_any_value(&AnyValue::String("test")),
            Some(Cell::Text("test".to_string()))
        );
        assert_eq!(
            Cell::from_any_value(&AnyValue::Boolean(false)),
            Some(Cell::Boolean(false))
        );
        assert_eq!(Cell::from_any_value(&AnyValue::Null), None);
    }
}
