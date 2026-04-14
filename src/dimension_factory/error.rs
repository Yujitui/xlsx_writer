//! Dimension Factory 错误类型定义

use std::fmt;

/// Dimension Factory 错误类型
#[derive(Debug)]
pub enum DimensionFactoryError {
    /// 列未找到
    ColumnNotFound(String),
    /// 索引越界
    IndexOutOfBounds(i32, usize),
    /// 类型不匹配
    TypeMismatch(String, String),
    /// Polars 错误
    PolarsError(polars::error::PolarsError),
    /// JSON 解析错误
    JsonError(serde_json::Error),
}

impl fmt::Display for DimensionFactoryError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            DimensionFactoryError::ColumnNotFound(col) => {
                write!(f, "Column not found: {}", col)
            }
            DimensionFactoryError::IndexOutOfBounds(idx, max) => {
                write!(f, "Index {} out of bounds (max: {})", idx, max)
            }
            DimensionFactoryError::TypeMismatch(expected, found) => {
                write!(f, "Type mismatch: expected {}, found {}", expected, found)
            }
            DimensionFactoryError::PolarsError(e) => {
                write!(f, "Polars error: {}", e)
            }
            DimensionFactoryError::JsonError(e) => {
                write!(f, "JSON error: {}", e)
            }
        }
    }
}

impl std::error::Error for DimensionFactoryError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            DimensionFactoryError::PolarsError(e) => Some(e),
            DimensionFactoryError::JsonError(e) => Some(e),
            _ => None,
        }
    }
}

impl From<polars::error::PolarsError> for DimensionFactoryError {
    fn from(err: polars::error::PolarsError) -> Self {
        DimensionFactoryError::PolarsError(err)
    }
}

impl From<serde_json::Error> for DimensionFactoryError {
    fn from(err: serde_json::Error) -> Self {
        DimensionFactoryError::JsonError(err)
    }
}
