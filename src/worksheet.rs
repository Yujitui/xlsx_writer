use polars::prelude::DataFrame;
use std::collections::HashMap;
use crate::error::XlsxError;

/// Excel 工作表（Sheet）名稱中禁止出現的特殊字符。
///
/// 根據 Excel 規格，這些字符會導致文件損壞或無法保存。
/// 在構建新的 WorkSheet 任務時，必須對名稱進行此項校驗。
const FORBIDDEN_SHEET_NAME_CHARS: [char; 7] = ['\\', '/', '?', '*', ':', '[', ']'];

/// 表示單個 Excel 工作表的導出任務。
///
/// 封裝了數據內容、名稱以及針對特定坐標的樣式映射。
#[derive(Debug)]
pub struct WorkSheet {

    /// 待導出的 Polars DataFrame 數據源。
    pub df: DataFrame,
    /// 工作表在 Excel 中顯示的名稱。
    pub name: String,
    /// 單元格樣式映射。
    ///
    /// 鍵為 `(行, 列)` 坐標，值為預定義樣式池中的樣式名稱。
    /// 行號 `0` 對應表頭，`1..N` 對應數據行。
    pub style_map: Option<HashMap<(u32, u16), String>>,

}

impl WorkSheet {

    /// 創建並校驗一個新的 `WorkSheet` 實例。
    ///
    /// # 參數
    /// * `df` - 數據源。
    /// * `name` - 初始工作表名稱。
    /// * `style_map` - 可選的樣式坐標映射。
    ///
    /// # 錯誤 (Returns)
    /// * [`XlsxError::EmptyDataFrame`] - 如果數據表沒有任何行或列。
    /// * [`XlsxError::InvalidName`] - 如果名稱不符合 Excel 命名規範（長度、字符、空值）。
    pub fn new(df: DataFrame, name: String, style_map: Option<HashMap<(u32, u16), String>>) -> Result<Self, XlsxError> {
        // 1. 嚴格數據校驗：拒絕導出空表，確保物理意義上的導出可行性。
        if df.height() == 0 || df.width() == 0 { return Err(XlsxError::EmptyDataFrame); }

        // 2. 名稱規範性校驗：
        // - 移除前後空格。
        // - 檢查長度（Excel 限制為 31 個字符，非字節數）。
        // - 檢查是否包含系統禁用字符。
        let trimmed = name.trim();
        let is_invalid =
            trimmed.is_empty() || // 为空值
                trimmed.chars().count() > 31 || // 超过31个字符
                trimmed.contains(&FORBIDDEN_SHEET_NAME_CHARS[..]); // 含有非法字符
        if is_invalid { return Err(XlsxError::InvalidName(name)); }

        // 3. 樣式坐標清洗（容錯性設計）：
        // 自動剔除超出當前數據表維度的無效坐標，以及空的樣式名。
        // 這樣可以防止在導出循環中出現下標越界或邏輯錯誤。
        let style_map = style_map.and_then(|mut map| {
            let max_row = df.height() as u32; // 行邊界（含表頭）
            let max_col = df.width() as u16;  // 列邊界索引

            // 过滤逻辑
            map.retain(|&(r, c), style_name| {
                r <= max_row && c < max_col && !style_name.trim().is_empty()
            });

            if map.is_empty() { None } else { Some(map) }
        });

        Ok(WorkSheet {df, name, style_map})
    }

}