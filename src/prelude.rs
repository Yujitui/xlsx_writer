//! 庫的預導入模組：包含導出 Excel 所需的核心類型與 Trait。

// 1. 核心數據結構
pub use crate::workbook::Workbook;
pub use crate::worksheet::WorkSheet;
pub use crate::readsheet::ReadSheet;

pub use crate::xls::XlsWorkbook;

// 2. 樣式工廠引擎
pub use crate::style_factory::StyleFactory;
pub use crate::style_factory::rules::{StyleRule, ApplyAction, StyleOverride};
pub use crate::style_factory::condition::StyleCondition;
pub use crate::style_library::{StyleLibrary, StyleDefinition};
pub use crate::merge_factory::MergeFactory;

// 3. 錯誤體系
pub use crate::style_factory::error::StyleFactoryError;

// 4. 外部依賴的關鍵類型（方便用戶配置樣式）
pub use rust_xlsxwriter::{Format, Color, FormatBorder, FormatAlign};
pub use polars::prelude::DataFrame;