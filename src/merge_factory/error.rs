use thiserror::Error;
use polars::prelude::PolarsError;

#[derive(Error, Debug)]
pub enum MergeFactoryError {
    /// 規則中指定的 targets 列名在當前 DataFrame 中不存在。
    #[error("合併規則引用了不存在的列: {0}")]
    ColumnNotFound(String),

    /// 索引越界（針對 Fixed 模式）。
    #[error("合併行索引 {0} 超出範圍 (高度: {1})")]
    IndexOutOfBounds(u32, usize),

    /// 封裝 Polars 內部的錯誤。
    #[error("Polars 運算異常: {0}")]
    PolarsError(#[from] PolarsError),

    /// 預留：合併區間重疊（未來擴展）。
    #[error("合併區間衝突: ({0},{1}) 到 ({2},{3})")]
    RangeConflict(u32, u16, u32, u16),
}