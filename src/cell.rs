//! 通用 Excel 单元格类型
//!
//! 本模块定义了 Excel 工作表中通用的单元格数据类型，
//! 可被 xlsx 和 xls 格式共用。

use polars::datatypes::AnyValue;
use rust_xlsxwriter::Format;

/// 富文本的一个片段（run）。
///
/// 每个片段由一段文本及其字体格式组成。多个片段按顺序拼接即构成
/// 一个富文本字符串，可用于单元格内的部分格式化（如某段加粗、变色）。
#[derive(Debug, Clone, PartialEq)]
pub struct RichTextSegment {
    /// 片段文本内容（不可为空字符串）
    pub text: String,
    /// 该片段的字体格式
    pub format: Format,
}

/// 通用 Excel 单元格类型
///
/// 封装了 Excel 单元格可以包含的不同数据类型。
/// 这是一个简化版本，专注于数据导出场景，暂不支持公式。
#[derive(Debug, Clone, PartialEq)]
pub enum Cell {
    /// 数值类型（f64）
    ///
    /// Excel 内部将数字存储为 64 位浮点数。
    /// 这包括整数、小数以及日期（以序列号形式存储）。
    Number(f64),

    /// 文本类型
    ///
    /// 纯文本内容。
    Text(String),

    /// 富文本类型
    ///
    /// 由多个带格式的文本片段组成，支持单元格内部分格式化。
    /// 注意：仅 .xlsx 写入路径支持完整格式；.xls 路径会降级为纯文本。
    RichText(Vec<RichTextSegment>),

    /// 布尔类型
    ///
    /// 逻辑值 true 或 false。
    Boolean(bool),
}

/// 将有限浮点数转为 Cell::Number，非有限值（NaN/Inf/-Inf）视为空值
fn finite_or_none(num: f64) -> Option<Cell> {
    if num.is_finite() {
        Some(Cell::Number(num))
    } else {
        None
    }
}

impl Cell {
    /// 将单元格内容转换为字符串
    pub fn to_string(&self) -> String {
        match self {
            Cell::Number(n) => n.to_string(),
            Cell::Text(s) => s.clone(),
            Cell::RichText(segments) => segments
                .iter()
                .map(|s| s.text.clone())
                .collect::<String>(),
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
            AnyValue::Float64(n) => finite_or_none(*n),
            AnyValue::Float32(n) => finite_or_none(*n as f64),
            AnyValue::Int8(v) => finite_or_none(*v as f64),
            AnyValue::Int16(v) => finite_or_none(*v as f64),
            AnyValue::Int32(v) => finite_or_none(*v as f64),
            AnyValue::Int64(v) => finite_or_none(*v as f64),
            AnyValue::UInt8(v) => finite_or_none(*v as f64),
            AnyValue::UInt16(v) => finite_or_none(*v as f64),
            AnyValue::UInt32(v) => finite_or_none(*v as f64),
            AnyValue::UInt64(v) => finite_or_none(*v as f64),
            AnyValue::Boolean(b) => Some(Cell::Boolean(*b)),
            AnyValue::Date(days) => finite_or_none(*days as f64),
            AnyValue::Datetime(v, unit, _) => {
                // 转换为 Excel 日期序列号（从 1899-12-30 起的天数）
                let days = match unit {
                    polars::datatypes::TimeUnit::Milliseconds => *v as f64 / 86_400_000.0,
                    polars::datatypes::TimeUnit::Microseconds => *v as f64 / 86_400_000_000.0,
                    polars::datatypes::TimeUnit::Nanoseconds => *v as f64 / 86_400_000_000_000.0,
                };
                finite_or_none(days)
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
    fn test_rich_text_to_string_concatenates() {
        let bold = Format::new().set_bold();
        let red = Format::new().set_font_color(rust_xlsxwriter::Color::Red);
        let rich = Cell::RichText(vec![
            RichTextSegment {
                text: "Hello".to_string(),
                format: bold,
            },
            RichTextSegment {
                text: " World".to_string(),
                format: red,
            },
        ]);
        assert_eq!(rich.to_string(), "Hello World");
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

    #[test]
    fn test_from_any_value_nan_inf_returns_none() {
        assert_eq!(Cell::from_any_value(&AnyValue::Float64(f64::NAN)), None);
        assert_eq!(
            Cell::from_any_value(&AnyValue::Float64(f64::INFINITY)),
            None
        );
        assert_eq!(
            Cell::from_any_value(&AnyValue::Float64(f64::NEG_INFINITY)),
            None
        );
        assert_eq!(Cell::from_any_value(&AnyValue::Float32(f32::NAN)), None);
        assert_eq!(
            Cell::from_any_value(&AnyValue::Float32(f32::INFINITY)),
            None
        );
    }

    #[test]
    fn test_finite_or_none() {
        assert_eq!(finite_or_none(1.0), Some(Cell::Number(1.0)));
        assert_eq!(finite_or_none(f64::NAN), None);
        assert_eq!(finite_or_none(f64::INFINITY), None);
        assert_eq!(finite_or_none(f64::NEG_INFINITY), None);
    }
}
