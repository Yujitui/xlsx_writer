pub mod cell;
mod error;
pub mod merge_factory;
pub mod prelude;
pub mod readsheet;
pub mod sheet_region;
pub mod style_factory;
pub mod style_library;
mod workbook;
mod worksheet;
pub mod xls_records;

// 常用组件提升到顶层
pub use cell::Cell;
pub use error::{XlsError, XlsxError};
pub use sheet_region::{RegionType, SheetRegion};
pub use workbook::Workbook;
pub use worksheet::WorkSheet;
pub use xls_records::RecordType;
