use serde::Deserialize;
use crate::style_factory::Condition;

/// 局部樣式覆蓋規則。
///
/// 用於在已選中行（Matched Row）的內部，針對符合特定條件的單元格進行樣式替換。
#[derive(Deserialize, Debug)]
pub struct StyleOverride {

    /// 當列判定條件（`col_conditions`）滿足時，該單元格應用的新樣式名。
    ///
    /// 此樣式會覆蓋所在行的 `ApplyAction::style`（基礎樣式）。
    pub style: String,

    /// 決定行內哪些列或單元格命中此覆蓋規則的條件集。
    ///
    /// 複用了 [`Condition`] 系統。通常用於：
    /// 1. 指定特定列名（`Match` 類型）。
    /// 2. 根據該格數據值判定（`ValueRange` 類型）。
    /// 3. 檢查內容包含（`Find` 類型）。
    pub col_conditions: Vec<Condition>,

}

/// 定義命中行規則後的樣式套用動作。
///
/// 採用兩層渲染邏輯：首先為整行應用基礎樣式，隨後根據 `overrides` 進行局部修正。
#[derive(Deserialize, Debug)]
pub struct ApplyAction {

    /// 命中的整行保底樣式名稱。
    ///
    /// 該樣式名必須存在於 `Workbook` 的樣式池中。若命中，該行所有單元格均先套用此樣式。
    pub style: String,

    /// 可選的局部樣式覆蓋規則列表。
    ///
    /// 用於在已選中的行內，針對特定列或滿足特定條件的單元格進行樣式替換。
    /// 例如：整行背景變灰，但「金額」列若為負數則顯示紅字。
    pub overrides: Option<Vec<StyleOverride>>,

}

/// 定義單條樣式生成規則。
///
/// 每一條 `StyleRule` 都是一個獨立的邏輯塊，描述了在特定數據條件下應如何著色。
#[derive(Deserialize, Debug)]
pub struct StyleRule {

    /// 行過濾器集合。
    ///
    /// 這是一個邏輯「與」（AND）關係的列表。只有當 DataFrame 中的某一行
    /// 同時滿足此列表中定義的所有 [`Condition`] 時，該行才會被視為「命中」。
    pub row_conditions: Vec<Condition>,

    /// 命中後的樣式套用策略。
    ///
    /// 包含整行的基礎樣式配置以及可選的單元格級別覆蓋邏輯。
    pub apply: ApplyAction,

}