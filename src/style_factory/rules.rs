use crate::style_factory::StyleCondition;
use serde::Deserialize;

/// 局部樣式覆蓋規則
///
/// 提供細粒度的樣式控制，在已選中行的內部針對特定單元格進行樣式替換。
#[derive(Deserialize, Debug)]
pub struct StyleOverride {
    /// 覆蓋樣式名稱
    ///
    /// 當列條件滿足時應用的樣式，會覆蓋所在行的基礎樣式。
    /// 該樣式必須在 Workbook 的樣式池中預先定義。
    pub style: String,

    /// 列條件集合
    ///
    /// 定義哪些單元格應該應用此覆蓋樣式。支持多種條件類型：
    /// - `Match`: 指定特定列名
    /// - `ValueRange`: 基於數值範圍的條件
    /// - `Find`: 字符串內容匹配
    /// - `Equal`: 列間一致性檢查
    ///
    /// # 邏輯關係
    /// 多個條件之間為邏輯「與」（AND）關係。
    pub col_conditions: Vec<StyleCondition>,
}

/// 樣式應用策略
///
/// 定義命中行規則後的兩層樣式應用機制。
#[derive(Deserialize, Debug)]
pub struct ApplyAction {
    /// 行級基礎樣式
    ///
    /// 命中行的所有單元格首先應用此樣式。
    /// 必須在 Workbook 的樣式池中預先定義。
    pub style: String,

    /// 單元格級覆蓋規則
    ///
    /// 可選的局部樣式覆蓋，用於精確控制特定單元格的顯示效果。
    pub overrides: Option<Vec<StyleOverride>>,
}

/// 樣式規則定義
///
/// 描述在特定數據條件下如何應用樣式的完整規則。
#[derive(Deserialize, Debug)]
pub struct StyleRule {
    /// 行選擇條件
    ///
    /// 定義哪些行應該應用此樣式規則。多個條件之間為邏輯「與」關係。
    /// 只有當一行同時滿足所有條件時才會被選中。
    pub row_conditions: Vec<StyleCondition>,

    /// 樣式應用策略
    ///
    /// 定義選中行後的樣式處理方式，包括基礎樣式和可選的覆蓋規則。
    pub apply: ApplyAction,
}
