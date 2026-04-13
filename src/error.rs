use thiserror::Error;

/// XLS 格式相关的错误类型
#[derive(Error, Debug)]
pub enum XlsError {
    #[error("IO error: {0}")]
    IoError(#[from] std::io::Error),

    #[error("Invalid file format: {0}")]
    InvalidFormat(String),

    #[error("String decoding failed: {0}")]
    Encoding(String),

    #[error("Cell index out of bounds: row={0}, col={1}")]
    OutOfBounds(usize, usize),

    #[error("Unexpected end of file")]
    UnexpectedEof,
}

/// 導出 Excel 過程中的專用錯誤枚舉。
///
/// 該枚舉涵蓋了從數據源校驗、業務邏輯衝突到物理磁盤 IO 的全鏈路異常。
#[derive(Error, Debug)]
pub enum XlsxError {
    /// Excel 文件为空或不含任何工作表
    #[error("Excel 文件未包含任何工作表")]
    NoSheetsFound,

    /// 指定的工作表不存在
    #[error("找不到名为 '{0}' 的工作表")]
    SheetNotFound(String),

    /// 无法读取标题行（通常是第一行缺失或为空）
    #[error("无法提取标题行，可能是由于表格为空或首行为全空")]
    MissingHeaderRow,

    #[error("列 '{column}' 类型转换失败，即使已尝试转为 String")]
    ConversionFailed { column: String },

    /// 文件本身存在问题（如加密、损坏）
    #[error("Excel 解析错误: {0}")]
    CalamineError(#[from] calamine::XlsxError),

    /// 當傳入的 Polars DataFrame 沒有任何行或列時觸發。
    #[error("DataFrame 內容為空，無法導出")]
    EmptyDataFrame,

    /// 工作表名稱不合法時觸發。
    ///
    /// 不合法的定義包括：空字符串、長度超過 31 個字符、或包含禁用字符。
    #[error("工作表名稱 '{0}' 不可用")]
    InvalidName(String),

    /// 在同一個 Workbook 中嘗試插入多個同名工作表時觸發。
    #[error("重複的工作表名稱: {0}")]
    DuplicateName(String),

    /// 當單元格指定的樣式標籤（String）在 Exporter 的樣式緩存池中找不到時觸發。
    #[error("樣式名 '{0}' 未在樣式池中定義")]
    UnknownStyle(std::sync::Arc<str>),

    /// 封裝標準庫的 IO 錯誤（如路徑無權限、磁盤空間不足）。
    ///
    /// 通過 `#[from]` 支持使用 `?` 運算符自動轉換。
    #[error("IO 錯誤: {0}")]
    IoError(#[from] std::io::Error),

    /// 封裝 `rust_xlsxwriter` 引擎內部的特定錯誤。
    ///
    /// 用於捕獲 Excel 格式限制或引擎運算異常。
    #[error("Excel 導出錯誤: {0}")]
    XlsxError(#[from] rust_xlsxwriter::XlsxError),

    #[error("樣式引擎錯誤: {0}")]
    StyleError(#[from] crate::style_factory::StyleFactoryError),

    #[error("合併引擎錯誤: {0}")]
    MergeError(#[from] crate::merge_factory::MergeFactoryError),

    /// 通用的字符串錯誤訊息
    #[error("{0}")]
    GenericError(String),
}
