use std::collections::HashMap;
use std::sync::Arc;
use polars::datatypes::DataType;
use polars::prelude::*;
use crate::style_factory::{parse_criteria, Condition, StyleFactoryError, StyleOverride, StyleRule};

/// 樣式生成工廠。
///
/// 作為導出流程中的中間層，它負責根據預定義的邏輯規則（Rules）
/// 對 DataFrame 進行掃描，並計算出每個單元格應採用的樣式標籤。
pub struct StyleFactory {

    /// 規則集。
    ///
    /// 存儲了從配置文件中加載的所有 [`StyleRule`] 實例。
    /// 執行時會按照此 Vec 的索引順序依次處理，後面的規則擁有更高的覆蓋優先級。
    pub rules: Vec<StyleRule>,

}

impl StyleFactory {

    /// 從 `serde_json::Value` 對象構造 `StyleFactory` 實例。
    ///
    /// 該方法專注於數據的層級提取與類型轉換，是實現動態配置加載的核心。
    ///
    /// # 參數
    /// * `data` - 已解析的 JSON 數組或對象。預期包含一個名為 `"rules"` 的鍵。
    ///
    /// # 邏輯流程
    /// 1. **路徑提取**：從輸入的 JSON 根節點定位到 `"rules"` 數組。
    /// 2. **強類型轉換**：利用 `serde_json::from_value` 將 JSON 片段轉化為 `Vec<StyleRule>`。
    /// 3. **狀態檢查**：如果規則集為空，向標準錯誤流輸出警告，但仍允許構建成功（支持空導出模式）。
    ///
    /// # 錯誤 (Returns)
    /// * `Box<dyn std::error::Error>` - 如果 JSON 結構不符合 [`StyleRule`] 的定義，則拋出反序列化異常。
    pub fn new(data: serde_json::Value) -> Result<Self, Box<dyn std::error::Error>> {
        // 從 JSON 根對象中提取 "rules" 鍵。
        // 若 "rules" 不存在，serde_json 會返回 Null 值，隨後的反序列化將捕獲此異常。
        let rules: Vec<StyleRule> = serde_json::from_value(data["rules"].clone())?;

        // 防禦性檢查：空規則集雖然技術上合法，但通常意味著配置文件加載異常。
        if rules.is_empty() {
            eprintln!("Warning: StyleFactory initialized with empty rules.");
        }

        Ok(Self { rules })
    }

    /// 從 JSON 格式的字符串構造 `StyleFactory` 實例。
    ///
    /// 該方法封裝了反序列化的初步過程，是應用程序讀取外部配置文件的標準路徑。
    ///
    /// # 參數
    /// * `json_str` - 包含樣式規則定義的原始 JSON 字符串。
    ///
    /// # 邏輯流程
    /// 1. **初步解析**：將字符串轉換為弱類型的 [`serde_json::Value`]。
    /// 2. **邏輯銜接**：將解析後的 Value 對象傳遞給 [`Self::new`] 進行規則提取。
    ///
    /// # 錯誤 (Returns)
    /// * `Box<dyn std::error::Error>` - 如果字符串不符合 JSON 語法，或內部規則結構校驗失敗。
    pub fn from_json_str(json_str: &str) -> Result<Self, Box<dyn std::error::Error>> {
        // 1. 執行字符串級別的 JSON 反序列化
        // 若語法錯誤（如缺少括號），將在此處透過 `?` 提前返回 Err
        let value: serde_json::Value = serde_json::from_str(json_str)?;

        // 2. 調用核心構造函數完成對象構建
        Self::new(value)
    }

    /// 評估一組條件，生成一個布爾遮罩（Boolean Mask）。
    ///
    /// 該方法遍歷所有條件並執行邏輯「與」（AND）運算。只有滿足所有條件的行，在最終遮罩中才為 `true`。
    ///
    /// # 參數
    /// * `df` - 指向待掃描的 DataFrame 的引用。
    /// * `conditions` - 待評估的條件切片。
    ///
    /// # 錯誤處理 (Errors)
    /// * 返回 [`StyleFactoryError`]，包含列名未找到、類型不匹配或 Polars 運算異常。
    fn evaluate_conditions(&self, df: &DataFrame, conditions: &[Condition]) -> Result<BooleanChunked, StyleFactoryError> {
        let height = df.height();
        // 1. 初始化全 true 遮罩：作為邏輯「與」運算的起點
        let mut full_mask = BooleanChunked::full("mask".into(), true, height);

        for cond in conditions {
            // 2. 針對每種條件類型執行特定的向量化判定
            let mask = match cond {
                // A. 索引定位：手動構建位圖（Bitmap），支持負數偏移
                Condition::Index { criteria, .. } => {
                    let mut mask_vec = vec![false; height];
                    for &phys_idx in criteria {
                        // 處理負數：-1 代表最後一行數據 (height)
                        if phys_idx == 0 { continue; }

                        // 物理 1..N 对应 DF 的 0..N-1
                        let df_idx = (phys_idx - 1) as usize;
                        if df_idx < height { mask_vec[df_idx] = true; }
                        else {
                            // 嚴格模式：索引越界直接拋錯
                            return Err(StyleFactoryError::IndexOutOfBounds(phys_idx, height));
                        }
                    }
                    Ok::<BooleanChunked, StyleFactoryError>(BooleanChunked::from_slice("mask".into(), &mask_vec))
                },
                // B. 數值範圍：解析算子（如 ">"）並執行 Series 比較
                Condition::ValueRange { targets, criteria } => {
                    let (op, val) = parse_criteria(criteria);
                    let mut m = BooleanChunked::full("mask".into(), true, height);
                    for col_name in targets {
                        let col = df.column(col_name)
                            .map_err(|_| StyleFactoryError::ColumnNotFound(col_name.clone()))?;
                        let series = col.as_materialized_series();

                        // 防禦性檢查：數值比較僅適用於數值類型列
                        if !series.dtype().is_numeric() {
                            return Err(StyleFactoryError::TypeMismatch(col_name.clone(), format!("{:?}", series.dtype())));
                        }

                        // 構建與列類型一致的比較基準值（Literal）
                        let val_lit = Series::new("lit".into(), vec![val]).cast(series.dtype())?;

                        let col_mask = match op {
                            ">" => series.gt(&val_lit)?,
                            "<" => series.lt(&val_lit)?,
                            ">=" => series.gt_eq(&val_lit)?,
                            "<=" => series.lt_eq(&val_lit)?,
                            _ => series.equal(&val_lit)?,
                        };
                        m = m & col_mask;
                    }
                    Ok::<BooleanChunked, StyleFactoryError>(m)
                },
                // C. 集合匹配：利用 HashSet 實現 O(1) 的成員檢查
                Condition::Match { targets, criteria } => {
                    let mut m = BooleanChunked::full("mask".into(), true, height);
                    // 為了極致性能與穩定性，將基準列表轉為 HashSet
                    // 這樣在循環中查詢的時間複雜度是 O(1)
                    let criteria_set: std::collections::HashSet<&str> = criteria.iter().map(|s| s.as_str()).collect();

                    for col_name in targets {
                        let col = df.column(col_name)
                            .map_err(|_| StyleFactoryError::ColumnNotFound(col_name.clone()))?;
                        let series = col.as_materialized_series();

                        // 透過迭代器手動構建遮罩，徹底避開 Series.is_in() 的 Trait 問題
                        // 這對於 String 類型列最為直接
                        let col_mask: BooleanChunked = if series.dtype() == &DataType::String {
                            // 字符串列：直接利用 .str() 命名空間迭代
                            series.str()
                                .map_err(|e| StyleFactoryError::PolarsError(e))?
                                .into_iter()
                                .map(|opt_val| {
                                    opt_val.map(|v| criteria_set.contains(v)).unwrap_or(false)
                                })
                                .collect()
                        } else {
                            // 非字符串列：將數值轉為字符串標籤後匹配
                            series.iter().map(|val| {
                                let s = format!("{}", val);
                                let s_clean = s.trim_matches('\"');
                                criteria_set.contains(s_clean)
                            }).collect()
                        };
                        // 5. 合併遮罩
                        m = m & col_mask;
                    }
                    Ok::<BooleanChunked, StyleFactoryError>(m)
                },
                // D. 字符串查找：調用 Polars 內建的 .str().contains()
                Condition::Find { targets, criteria } => {
                    // 1. 初始化全為 true 的遮罩
                    let mut m = BooleanChunked::full("mask".into(), true, height);

                    for col_name in targets {
                        // 2. 獲取 Column 並轉為 Series
                        let col = df.column(col_name)
                            .map_err(|_| StyleFactoryError::ColumnNotFound(col_name.clone()))?;
                        let series = col.as_materialized_series();

                        // 3. 穩健校驗：確保該列是字符串類型
                        if series.dtype() != &DataType::String {
                            return Err(StyleFactoryError::TypeMismatch(col_name.clone(), format!("{:?}", series.dtype())));
                        }

                        // 4. 調用 .str() 命名空間的 contains 方法
                        // 參數 1: 子串 (criteria)
                        // 參數 2: 是否使用正則表達式 (我們設為 false，僅進行簡單子串查找)
                        let col_mask = series
                            .str()
                            .map_err(|e| StyleFactoryError::PolarsError(e))?
                            .contains(criteria, false)
                            .map_err(|e| StyleFactoryError::PolarsError(e))?;

                        // 5. 合併遮罩
                        m = m & col_mask;
                    }
                    Ok::<BooleanChunked, StyleFactoryError>(m)
                },
                // E. 列間相等：多列對等橫向比較
                Condition::Equal { targets, criteria } => {
                    let mut m = BooleanChunked::full("mask".into(), true, height);
                    if targets.len() >= 2 {
                        let first_series = df.column(&targets[0])
                            .map_err(|_| StyleFactoryError::ColumnNotFound(targets[0].clone()))?
                            .as_materialized_series();
                        for col_name in &targets[1..] {
                            let other_series = df.column(col_name)
                                .map_err(|_| StyleFactoryError::ColumnNotFound(col_name.clone()))?
                                .as_materialized_series();
                            let eq_mask = first_series
                                .equal(other_series)
                                .map_err(|e| StyleFactoryError::PolarsError(e))?;
                            m = m & eq_mask;
                        }
                    }
                    // 這裡判斷 bool 標籤決定是否取反
                    if *criteria { Ok::<BooleanChunked, StyleFactoryError>(m) } else { Ok(!m) }
                },
                // F. 索引排除：手動構建位圖（Bitmap），支持負數偏移
                Condition::ExcludeRows { criteria, .. } => {
                    let mut mask_vec = vec![true; height];
                    let [p_start, p_end] = *criteria;

                    // 将物理区间转换为 DF 逻辑区间
                    // 例子：排除物理 [1, 1] (第一行数据) -> 对应 DF [0, 0]
                    let df_start = if p_start <= 0 { 0 } else { (p_start - 1) as usize };
                    let df_end = if p_end <= 0 {
                        if p_end == 0 { 0 } else { (height as i32 + p_end) as usize } // 处理负数
                    } else {
                        (p_end - 1) as usize
                    };

                    for i in df_start..=df_end {
                        if i < height { mask_vec[i] = false; }
                    }
                    Ok(BooleanChunked::from_slice("mask".into(), &mask_vec))
                },
            }?; // 透過 ? 解開分支 Result 並進行錯誤轉換
            full_mask = full_mask & mask;
        }

        Ok(full_mask)
    }

    /// 統一的列定位引擎：適用於標題行 Override 和 數據行 Override。
    ///
    /// 它根據列名或索引返回命中列的物理清單。
    fn evaluate_col_indices(&self, df: &DataFrame, conditions: &[Condition]) -> Result<Vec<u16>, StyleFactoryError> {
        // 1. 嚴格檢查：禁止數據驅動型算子出現在列定義中
        let has_data_op = conditions.iter().any(|c| matches!(c, Condition::ValueRange{..} | Condition::Equal{..}));
        if has_data_op {
            return Err(StyleFactoryError::TypeMismatch("列定位".into(), "數據算子".into()));
        }

        // 2. 獲取所有列名供匹配
        let column_names = df.get_column_names();
        let mut matched = std::collections::HashSet::new();

        for cond in conditions {
            match cond {
                // 通過列名精確匹配 (Match) 或 關鍵字 (Find)
                c if c.get_targets().is_some() => {
                    if let Some(targets) = c.get_targets() {
                        for name in targets {
                            if let Some(idx) = df.get_column_index(name) {
                                matched.insert(idx as u16);
                            }
                        }
                    }
                },
                // 通過物理位置匹配 (Index)
                Condition::Index { criteria, .. } => {
                    for &idx in criteria {
                        if idx >= 0 && (idx as usize) < column_names.len() {
                            matched.insert(idx as u16);
                        }
                    }
                },
                _ => {}
            }
        }
        Ok(matched.into_iter().collect())
    }

    /// 判定當前規則是否命中 Excel 的物理標題行（第 0 行）。
    ///
    /// 該函數是樣式工廠實現「物理座標感知」的核心組件，負責在數據掃描前過濾掉與表頭無關的規則。
    ///
    /// # 准入邏輯
    /// 1. **屬性隔離**：如果規則中包含任何「數據驅動型」條件（如 `ValueRange`, `Match`, `Find`, `Equal`），
    ///    則該規則被視為僅針對 DataFrame 數據內容，自動判定為「不命中」表頭。
    /// 2. **座標判定**：如果規則僅由「位置驅動型」條件（`Index` 或 `ExcludeRows`）組成，則執行物理座標匹配。
    ///
    /// # 匹配規則
    /// * **Index 命中**：`row_conditions` 中的 `Index` 條件必須顯式包含物理索引 `0`。
    /// * **Exclude 攔截**：如果 `ExcludeRows` 定義的排除區間 `[start, end]` 覆蓋了物理索引 `0`，則強制不命中。
    /// * **邏輯與 (AND)**：規則內的所有條件必須同時滿足對物理 `0` 號位的准入要求。
    ///
    /// # 參數
    /// * `rule` - 待評估的單條樣式生成規則。
    ///
    /// # 返回值 (Returns)
    /// * `true` - 該規則具備修改 Excel 物理第 0 行樣式的權限。
    /// * `false` - 該規則應被限制在數據行（Row 1..N）範圍內。
    fn rule_matches_header(&self, rule: &StyleRule) -> bool {
        // 如果规则中包含任何需要读取单元格内容的条件，则表头（无数据）不可能命中
        let has_content_dep = rule.row_conditions.iter().any(|cond| match cond {
            Condition::ValueRange { .. } | Condition::Match { .. } |
            Condition::Find { .. } | Condition::Equal { .. } => true,
            _ => false,
        });

        if has_content_dep { return false; }

        // 如果只有物理位置条件 (Index/ExcludeRows)，执行逻辑判定
        // 假设 criteria 里的 0 代表物理表头
        let mut hit = true;
        for cond in &rule.row_conditions {
            match cond {
                Condition::Index { criteria, .. } => {
                    if !criteria.contains(&0) { hit = false; }
                },
                Condition::ExcludeRows { criteria, .. } => {
                    // 如果排除范围包含 0，则不命中
                    if criteria[0] <= 0 && criteria[1] >= 0 { hit = false; }
                },
                _ => {}
            }
        }
        hit
    }

    /// 判定局部覆蓋規則（Override）是否適用於標題行。
    ///
    /// 在標題行場景下，允許基於列標識（Match/Find）或物理位置（Index）的判定，
    /// 但禁止基於數值邏輯（ValueRange/Equal）的判定。
    fn ov_matches_header(&self, ov: &StyleOverride) -> bool {
        !ov.col_conditions.iter().any(|cond| match cond {
            // 禁止：標題行沒有數值，無法執行範圍比較或多列相等檢查
            Condition::ValueRange { .. } | Condition::Equal { .. } => true,
            // 允許：Index (位置), Match (精確列名), Find (列名關鍵字)
            _ => false,
        })
    }

    /// 根据定义的样式规则（Rules）为 DataFrame 生成单元格样式映射。
    ///
    /// ### 工作原理
    /// 1. **行筛选**：遍历所有规则，首先评估 `row_conditions` 确定哪些行受该规则影响。
    /// 2. **基础样式填充**：对于命中的行，默认先将整行所有列应用 `rule.apply.style`。
    /// 3. **局部样式覆盖 (Overrides)**：
    ///    - 在命中的行内，进一步评估 `overrides` 中的列条件。
    ///    - 若特定单元格满足条件，则用 `ov.style` 覆盖该位置的基础样式。
    ///
    /// ### 参数
    /// * `df` - 输入的 Polars DataFrame，用于条件评估和索引查找。
    ///
    /// ### 返回值
    /// * `Ok(HashMap<(行索引, 列索引), 样式字符串>)` - 成功时返回样式映射表。
    /// * `Err` - 若规则配置无效（如行筛选器为空）或引用的列名不存在，则返回错误。
    ///
    /// ### 性能注意 (Performance Note)
    /// - **复杂度**：O(N_rules * N_rows * (N_cols + N_overrides))。
    /// - **内存**：会为每个受影响的单元格克隆样式字符串。对于超大型表格，建议优化样式存储（如使用 Arc）。
    /// - **优化建议**：目前在行循环内动态查找列索引，对于大表，建议预先缓存 `ColumnIndex`。
    pub fn execute(&self, df: &DataFrame) -> Result<HashMap<(u32, u16), Arc<str>>, Box<dyn std::error::Error>> {
        let mut style_map = HashMap::new();
        let width = df.width() as u16;

        for rule in &self.rules {
            // --- 【优化 1：使用 Arc 包装样式】 ---
            let base_style: Arc<str> = rule.apply.style.as_str().into();

            // --- 步驟 A：判定物理第 0 行 (表頭) ---
            // 只有不含數據判定（ValueRange/Match等）且满足 Index/Exclude 逻辑的规则能命中表头
            if self.rule_matches_header(rule) {
                // 1. 寫入標題基礎樣式
                for c in 0..width {
                    style_map.insert((0, c), Arc::clone(&base_style));
                }
                // 2. 處理標頭 Overrides
                if let Some(ref overrides) = rule.apply.overrides {
                    for ov in overrides {
                        if self.ov_matches_header(ov) {
                            let ov_style: Arc<str> = ov.style.as_str().into();
                            // 定位哪些列命中
                            let target_cols = self.evaluate_col_indices(df, &ov.col_conditions)?;
                            for c_idx in target_cols {
                                style_map.insert((0, c_idx), Arc::clone(&ov_style));
                            }
                        }
                    }
                }
            }

            // 1. 获取行掩码：确定当前规则作用于哪些行 (返回长度为 height 的布尔序列)
            let row_mask = self.evaluate_conditions(df, &rule.row_conditions)?;

            // --- 【优化 2：预处理 Overrides，避免在行循环内重复查找和计算】 ---
            let mut prepared_overrides = Vec::new();
            if let Some(ref overrides) = rule.apply.overrides {
                for ov in overrides {
                    // 使用統一的列定位引擎獲取物理列清單
                    let target_cols = self.evaluate_col_indices(df, &ov.col_conditions)?;
                    let ov_style: Arc<str> = ov.style.as_str().into();
                    prepared_overrides.push((target_cols, ov_style));
                }
            }

            for (r_idx, matched) in row_mask.into_iter().enumerate() {
                // 只处理匹配成功的行 (matched 为 Some(true))
                if let Some(true) = matched {
                    let r_phys = (r_idx + 1) as u32;
                    // 2. 应用整行基础样式：初始化该行所有列的默认风格
                    for c in 0..width {
                        // --- 【优化 4：Arc clone 仅增加引用计数，不分配内存】 ---
                        style_map.insert((r_phys, c), Arc::clone(&base_style));
                    }
                    // 3. 处理 Overrides：针对特定列或单元格进行样式覆盖
                    // 應用預存的列覆蓋 (極其高效：直接遍歷索引 Vec)
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