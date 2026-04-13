use thiserror::Error;

/// 樣式工廠引擎專用錯誤體系。
///
/// 涵蓋了從 JSON 解析、邏輯校驗到 Polars 向量化運算的完整錯誤鏈條。
#[derive(Error, Debug)]
pub enum StyleFactoryError {
    /// 樣式配置文件（JSON）語法錯誤或結構不符合 Schema 要求。
    ///
    /// 由 `serde_json` 拋出，包含具體的行號與列號信息。
    #[error("樣式配置文件解析失敗: {0}")]
    ConfigParseError(#[from] serde_json::Error),

    /// 規則中指定的 `targets` 列名在當前 DataFrame 中不存在。
    ///
    /// 常用於捕獲 JSON 配置中的拼寫錯誤。
    #[error("規則中引用了不存在的列: {0}")]
    ColumnNotFound(String),

    /// 數據類型與判斷算子衝突。
    ///
    /// 例如：嘗試對 `DataType::String` 類型的列執行 `ValueRange`（數值範圍）比較。
    /// 參數：(列名, 實際數據類型)。
    #[error("列 '{0}' 的數據類型 ({1}) 與判斷條件不匹配")]
    TypeMismatch(String, String),

    /// 顯式行索引定位超出了數據表的物理邊界。
    ///
    /// 在執行 `Index` 類型的條件判定，且處理完負數偏移後，索引仍不在 `0..height` 區間內時觸發。
    #[error("行索引 {0} 超出 DataFrame 範圍 (高度: {1})")]
    IndexOutOfBounds(i32, usize),

    /// 封裝 Polars 引擎在執行向量化比較（如 gt, lt, is_in）時產生的內部異常。
    #[error("Polars 運算異常: {0}")]
    PolarsError(#[from] polars::prelude::PolarsError),

    /// 規則中引用的樣式標籤（Label）未在 XlsxExporter 的樣式池中註冊。
    ///
    /// 確保樣式定義與樣式應用之間的標識符一致性。
    #[error("未定義的樣式標籤: {0}")]
    UnknownStyle(String),
}
