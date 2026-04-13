//! 庫的預導入模組：包含導出 Excel 所需的核心類型與 Trait。

// 1. 核心數據結構
pub use crate::readsheet::ReadSheet;
pub use crate::workbook::Workbook;
pub use crate::worksheet::WorkSheet;

// 2. 樣式工廠引擎
pub use crate::merge_factory::MergeFactory;
pub use crate::style_factory::condition::StyleCondition;
pub use crate::style_factory::rules::{ApplyAction, StyleOverride, StyleRule};
pub use crate::style_factory::StyleFactory;
pub use crate::style_library::{StyleDefinition, StyleLibrary};

// 3. 錯誤體系
pub use crate::style_factory::error::StyleFactoryError;

// 4. 外部依賴的關鍵類型（方便用戶配置樣式）
pub use polars::prelude::DataFrame;
pub use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder};
