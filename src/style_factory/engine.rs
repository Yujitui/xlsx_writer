use crate::style_factory::{evaluate_col_conditions, evaluate_row_conditions, StyleRule};
use polars::prelude::*;
use std::collections::HashMap;
use std::sync::Arc;

/// 樣式生成工廠
///
/// 作為 Excel 導出流程中的核心組件，負責根據預定義的邏輯規則
/// 對 DataFrame 進行掃描，並計算出每個單元格應採用的樣式標籤。
///
/// # 設計原則
///
/// ## 優先級機制
/// 規則按照 Vec 中的順序執行，索引越大的規則優先級越高，
/// 後面的規則可以覆蓋前面規則的結果。
///
/// ## 無狀態設計
/// 工廠本身不持有 DataFrame 的所有權，確保內存使用的高效性。
///
/// ## 可擴展架構
/// 通過 [`StyleRule`] 的多態設計，支持靈活的條件判斷和樣式應用。
pub struct StyleFactory {
    /// 規則集
    ///
    /// 存儲了從配置文件中加載的所有 [`StyleRule`] 實例。
    ///
    /// # 執行順序
    /// 按照 Vec 的索引順序依次處理：
    /// - 索引 0: 最低優先級
    /// - 索引 n: 最高優先級
    ///
    /// # 覆蓋機制
    /// 後面的規則會覆蓋前面規則對相同單元格的樣式設定。
    ///
    /// # 線程安全
    /// 規則集為只讀結構，可在多線程環境中安全共享。
    pub rules: Vec<StyleRule>,
}

impl StyleFactory {
    /// 創建新的樣式工廠實例
    pub fn new(value: serde_json::Value) -> Result<Self, Box<dyn std::error::Error>> {
        let rules: Vec<StyleRule> = if value.is_array() {
            serde_json::from_value(value)?
        } else {
            if !value.is_null() {
                eprintln!("Warning: StyleFactory input is not a valid array, using empty rules.");
            }
            vec![]
        };

        Ok(Self { rules })
    }

    /// 從 JSON 字符串創建樣式工廠實例
    pub fn from_json_str(json_str: &str) -> Result<Self, Box<dyn std::error::Error>> {
        let value: serde_json::Value = serde_json::from_str(json_str)?;
        Self::new(value)
    }

    /// 执行样式规则评估并生成单元格样式映射表
    pub fn execute(
        &self,
        df: &DataFrame,
    ) -> Result<HashMap<(u32, u16), Arc<str>>, Box<dyn std::error::Error>> {
        let mut style_map = HashMap::new();
        let width = df.width() as u16;

        for rule in &self.rules {
            let base_style: Arc<str> = rule.apply.style.as_str().into();
            let row_mask = evaluate_row_conditions(df, &rule.row_conditions)?;

            let mut prepared_overrides = Vec::new();
            if let Some(ref overrides) = rule.apply.overrides {
                for ov in overrides {
                    match evaluate_col_conditions(df, &ov.col_conditions) {
                        Ok(target_cols) => {
                            let ov_style: Arc<str> = ov.style.as_str().into();
                            prepared_overrides.push((target_cols, ov_style));
                        }
                        Err(_) => {
                            continue;
                        }
                    }
                }
            }

            for (r_idx, matched) in row_mask.into_iter().enumerate() {
                if let Some(true) = matched {
                    let r_phys = r_idx as u32;
                    for c in 0..width {
                        style_map.insert((r_phys, c), Arc::clone(&base_style));
                    }
                    for (indices, ov_style) in &prepared_overrides {
                        for &c_idx in indices {
                            style_map.insert((r_phys, c_idx), Arc::clone(ov_style));
                        }
                    }
                }
            }
        }
        Ok(style_map)
    }
}
