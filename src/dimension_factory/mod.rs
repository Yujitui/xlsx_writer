//! 维度工厂模块：负责根据规则计算行高和列宽

// 1. 声明内部子模块
pub mod engine;
pub mod error;
pub mod rule;

// 2. 重新导出核心类型
pub use engine::{DimensionFactory, DimensionResult};
pub use error::DimensionFactoryError;
pub use rule::{DimensionRule, DimensionTarget, DimensionValue};
