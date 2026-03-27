use serde::Deserialize;

/// 定義單元格或行的命中邏輯。
///
/// 採用 `snake_case` 命名規範（如 `value_range`），並透過 `type` 字段進行多態分發。
#[derive(Deserialize, Debug)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum Condition {

    /// 1. 物理位置定位。
    ///
    /// 依賴行號（Index）直接選定目標，不依賴數據內容。
    Index {
        /// 為了結構統一而保留，在 `Index` 模式下通常為空。
        #[serde(default)]
        targets: Vec<String>,
        /// 物理行索引列表。支持正數（從 0 開始）與負數（-1 代表最後一行）。
        criteria: Vec<i32>,
    },

    /// 2. 數值範圍判定。
    ///
    /// 檢查 `targets` 中指定的每一列數值是否滿足運算表達式。
    ValueRange {
        /// 目標列名列表。
        targets: Vec<String>,
        /// Excel风格的比較表達式字符串，如 `">5000"`, `"<100"`, `">=0.5"`。
        criteria: String,
    },

    /// 3. 集合成員匹配。
    ///
    /// 檢查 `targets` 中指定的列值是否包含在給定的白名單中。
    Match {
        /// 目標列名列表。
        targets: Vec<String>,
        /// 允許的值列表（字符串形式，匹配時會自動轉換類型）。
        criteria: Vec<String>,
    },

    /// 4. 字符串子串查找。
    ///
    /// 檢查 `targets` 中指定的字符串列是否包含特定片段。
    Find {
        /// 目標列名列表（僅限字符串類型列）。
        targets: Vec<String>,
        /// 待查找的子字符串。
        criteria: String,
    },

    /// 5. 列間一致性檢查。
    ///
    /// 比較 `targets` 列表中所有列的值是否彼此相等。
    Equal {
        /// 需要相互比較的列名列表（至少需要兩列）。
        targets: Vec<String>,
        /// 期望結果：`true` 代表要求全部相等，`false` 代表要求存在不相等。
        criteria: bool,
    },

    /// 6. 排除特定行范围（逻辑非）。
    ///
    /// 满足该范围内的行将被强制标记为「未命中」。
    ExcludeRows {
        /// 為了結構統一而保留，在 `ExcludeRows` 模式下通常為空。
        #[serde(default)]
        targets: Vec<String>,
        /// 范围起止：[开始, 结束] (包含起止)。
        /// 支持负数，例如 [0, 0] 代表排除第一行，[-1, -1] 代表排除最后一行。
        criteria: [i32; 2],
    },

}

impl Condition {

    /// 獲取當前條件所關聯的目標列名清單。
    ///
    /// 此方法用於在單元格級別的樣式覆蓋（`StyleOverride`）中確定塗色範圍。
    ///
    /// # 返回值 (Returns)
    /// * `Some(&Vec<String>)` - 對於基於內容的判斷（如 `ValueRange`, `Match` 等），返回其操作的目標列。
    /// * `None` - 對於 `Index` 類型，由於其僅依賴物理行號定位，不具備列屬性，故返回 `None`。
    ///
    /// # 設計意圖 (Design Intent)
    /// 通過返回引用而非所有權，避免了在遍歷大量規則時產生頻繁的內存分配與字符串克隆。
    pub fn get_targets(&self) -> Option<&Vec<String>> {
        match self {
            // Index\ExcludeRows 模式僅與物理行相關，執行引擎應據此判斷是否進行整行處理
            Condition::Index { .. } => None,
            Condition::ExcludeRows { .. } => None,

            // 所有基於內容的條件均指向一組具體的列
            Condition::ValueRange { targets, .. } => Some(targets),
            Condition::Match { targets, .. } => Some(targets),
            Condition::Find { targets, .. } => Some(targets),
            Condition::Equal { targets, .. } => Some(targets),
        }
    }

}