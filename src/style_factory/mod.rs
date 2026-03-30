//! 樣式工廠模組：負責將 JSON 規則轉換為 DataFrame 的單元格樣式映射。

// 1. 聲明內部子模組
pub mod condition;
pub mod rules;
pub mod engine;
pub mod error;

// 2. 重新導出核心類型（提高外部調用的便利性）
pub use condition::StyleCondition;
pub use rules::{StyleRule, ApplyAction, StyleOverride};
pub use engine::StyleFactory;
pub use error::StyleFactoryError;

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