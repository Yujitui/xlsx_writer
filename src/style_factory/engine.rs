use std::collections::HashMap;
use std::sync::Arc;
use polars::datatypes::DataType;
use polars::prelude::*;
use crate::style_factory::{parse_criteria, StyleCondition, StyleFactoryError, StyleRule};

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
    ///
    /// 從 JSON 配置值解析樣式規則，構建 StyleFactory 實例。
    ///
    /// # 參數
    ///
    /// * `value`: JSON 配置值，預期為包含樣式規則的數組
    ///   - 有效格式：`[{"row_conditions": [...], "apply": {...}}, ...]`
    ///   - 空數組：`[]` - 創建空規則集
    ///   - null 值：`null` - 創建空規則集
    ///   - 其他：觸發警告並使用空規則集
    ///
    /// # 返回值
    ///
    /// * `Ok(StyleFactory)` - 成功創建的工廠實例
    /// * `Err(Box<dyn std::error::Error>)` - JSON 解析失敗或其他錯誤
    pub fn new(value: serde_json::Value) -> Result<Self, Box<dyn std::error::Error>> {
        // 1. 執行輸入類型檢查與自動修正
        let rules: Vec<StyleRule> = if value.is_array() {
            // 如果是正常的數組，執行標準反序列化
            serde_json::from_value(value)?
        } else {
            // 如果是 Null 或其它無效類型，靜默降級為空
            if !value.is_null() {
                eprintln!("Warning: StyleFactory input is not a valid array, using empty rules.");
            }
            vec![]
        };

        Ok(Self { rules })
    }

    /// 從 JSON 字符串創建樣式工廠實例
    ///
    /// 提供從 JSON 配置字符串直接創建 StyleFactory 的便捷方法。
    ///
    /// # 參數
    ///
    /// * `json_str`: 包含樣式規則配置的 JSON 字符串
    ///   - 格式：有效的 JSON 數組字符串
    ///   - 示例：`"[{\"row_conditions\":[...],\"apply\":{...}}]"`
    ///
    /// # 返回值
    ///
    /// * `Ok(StyleFactory)` - 成功創建的工廠實例
    /// * `Err(Box<dyn std::error::Error>)` - 解析或構建過程中發生的錯誤
    pub fn from_json_str(json_str: &str) -> Result<Self, Box<dyn std::error::Error>> {
        // 1. 執行字符串級別的 JSON 反序列化
        // 若語法錯誤（如缺少括號），將在此處透過 `?` 提前返回 Err
        let value: serde_json::Value = serde_json::from_str(json_str)?;

        // 2. 調用核心構造函數完成對象構建
        Self::new(value)
    }

    /// 将列级别条件结果合并到行级别遮罩中
    ///
    /// 此函数专门用于将基于列数据的条件判断结果（如数值比较、字符串匹配等）
    /// 合并到完整的行级别样式遮罩中。通过在标题行位置插入 false 值，
    /// 确保标题行不会受到列级别条件的影响。
    ///
    /// # 参数说明
    ///
    /// * `mask`: 可变引用的目标行级别遮罩（长度 = height + 1）
    ///   - 索引 0: 标题行遮罩值
    ///   - 索引 1..height: 数据行遮罩值
    /// * `col_mask`: 列级别条件判断结果（长度 = height）
    ///   - 每个元素对应 DataFrame 中一行的条件判断结果
    fn merge_mask(mask: &mut BooleanChunked, col_mask: BooleanChunked) {
        // 使用 from_iter 直接构建，避免中间 Vec 分配
        // 构建过程：标题行(false) + 数据行(col_mask结果)
        let other = BooleanChunked::from_iter(
            std::iter::once(Some(false)) // 标题行不受列条件影响，始终为 false
                .chain(col_mask.into_iter()) // 数据行条件结果
        );

        // 更新原始 mask：逻辑"与"操作实现条件合并
        // 这确保了只有当原始遮罩和列条件都为 true 时，最终结果才为 true
        *mask = &*mask & &other;
    }

    /// 评估样式条件并生成匹配遮罩
    ///
    /// 对给定的条件集合执行向量化评估，返回标识哪些行匹配的布尔遮罩。
    /// 支持多种条件类型，使用逻辑"与"关系组合所有条件。
    ///
    /// # 参数说明
    ///
    /// * `df`: 输入的 DataFrame，提供数据源
    /// * `conditions`: 需要评估的条件列表
    ///
    /// # 返回值
    ///
    /// * `Ok(BooleanChunked)` - 评估结果遮罩，true 表示匹配
    /// * `Err(StyleFactoryError)` - 评估过程中发生的错误
    fn evaluate_conditions(&self, df: &DataFrame, conditions: &[StyleCondition]) -> Result<BooleanChunked, StyleFactoryError> {
        let height = df.height();
        // 1. 初始化全 true 遮罩：作為邏輯「與」運算的起點
        let mut full_mask = BooleanChunked::full("mask".into(), true, height + 1);

        for cond in conditions {
            // 2. 針對每種條件類型執行特定的向量化判定
            let mask = match cond {
                // A. 索引定位：手動構建位圖（Bitmap），支持負數偏移
                StyleCondition::Index { criteria, .. } => {
                    let mut mask = vec![false; height + 1];
                    for &phys_idx in criteria {
                        // 處理負數索引：-1 代表最後一行，-2 代表倒數第二行
                        let df_idx = if phys_idx < 0 {
                            // 负数索引转换：-1 -> height, -2 -> height-1
                            let adjusted_idx = height as i32 + phys_idx + 1;
                            if adjusted_idx < 0 {
                                return Err(StyleFactoryError::IndexOutOfBounds(phys_idx, height));
                            }
                            adjusted_idx as usize
                        } else {
                            // 正数索引转换：1 -> 1, 2 -> 1 (物理转逻辑)
                            phys_idx as usize
                        };

                        if df_idx < height {
                            mask[df_idx] = true;
                        } else {
                            // 嚴格模式：索引越界直接拋錯
                            return Err(StyleFactoryError::IndexOutOfBounds(phys_idx, height));
                        }
                    }
                    Ok::<BooleanChunked, StyleFactoryError>(BooleanChunked::from_slice("mask".into(), &mask))
                },
                // B. 數值範圍：解析算子（如 ">"）並執行 Series 比較
                StyleCondition::ValueRange { targets, criteria } => {
                    let (op, val) = parse_criteria(criteria);
                    // 设置标题行为 false
                    let mut mask = vec![true; height + 1];
                    mask[0] = false; // 标题行不参与数值比较
                    let mut mask = BooleanChunked::new("mask".into(), &mask);

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

                        Self::merge_mask(&mut mask, col_mask);
                    }
                    Ok::<BooleanChunked, StyleFactoryError>(mask)
                },
                // C. 集合匹配：利用 HashSet 實現 O(1) 的成員檢查
                StyleCondition::Match { targets, criteria } => {
                    // 设置标题行为 false
                    let mut mask = vec![true; height + 1];
                    mask[0] = false; // 标题行不参与数值比较
                    let mut mask = BooleanChunked::new("mask".into(), &mask);
                    // 為了極致性能與穩定性，將基準列表轉為 HashSet
                    // 這樣在循環中查詢的時間複雜度是 O(1)
                    let criteria_set: std::collections::HashSet<&str> = criteria.iter().map(|s| s.as_str()).collect();

                    for col_name in targets {
                        let col = df.column(col_name)
                            .map_err(|_| StyleFactoryError::ColumnNotFound(col_name.clone()))?;
                        let series = col.as_materialized_series();

                        // 透過迭代器手動構建遮罩，徹底避開 Series.is_in() 的 Trait 問題
                        // 這對於 String 類型列最為直接
                        let col_mask: Vec<bool> = if series.dtype() == &DataType::String {
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
                        // 将列匹配结果应用到最终结果（从索引1开始对应数据行）
                        Self::merge_mask(&mut mask, BooleanChunked::new("mask".into(), col_mask));
                    }
                    Ok::<BooleanChunked, StyleFactoryError>(mask)
                },
                // D. 字符串查找：調用 Polars 內建的 .str().contains()
                StyleCondition::Find { targets, criteria } => {
                    // 设置标题行为 false
                    let mut mask = vec![true; height + 1];
                    mask[0] = false; // 标题行不参与数值比较
                    let mut mask = BooleanChunked::new("mask".into(), &mask);

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

                        // 将列查找结果应用到最终结果（从索引1开始对应数据行）
                        Self::merge_mask(&mut mask, col_mask);
                    }
                    Ok::<BooleanChunked, StyleFactoryError>(mask)
                },
                // E. 列間相等：多列對等橫向比較
                StyleCondition::Equal { targets, criteria } => {
                    let mut mask = vec![true; height + 1];
                    mask[0] = false; // 标题行不参与数值比较
                    let mut mask = BooleanChunked::new("mask".into(), &mask);
                    if targets.len() >= 2 {
                        let first_series = df.column(&targets[0])
                            .map_err(|_| StyleFactoryError::ColumnNotFound(targets[0].clone()))?
                            .as_materialized_series();
                        for col_name in &targets[1..] {
                            let other_series = df.column(col_name)
                                .map_err(|_| StyleFactoryError::ColumnNotFound(col_name.clone()))?
                                .as_materialized_series();
                            let col_mask = first_series
                                .equal(other_series)
                                .map_err(|e| StyleFactoryError::PolarsError(e))?;
                            // 将列比较结果应用到最终结果（从索引1开始对应数据行）
                            Self::merge_mask(&mut mask, col_mask);
                        }
                    }
                    Ok::<BooleanChunked, StyleFactoryError>(if *criteria {mask} else { !mask })
                },
                // F. 索引排除：手動構建位圖（Bitmap），支持負數偏移
                StyleCondition::ExcludeRows { criteria, .. } => {
                    // 创建结果容器：长度为 height + 1，标题行(0)默认为 true
                    let mut mask = vec![true; height + 1];
                    // 标题行默认不排除（除非明确指定）
                    // mask[0] = true; // 标题行默认保留

                    let [p_start, p_end] = *criteria;

                    // 将物理区间转换为 Excel 行号区间
                    // 例子：排除物理 [1, 1] (第一行数据) -> 对应 Excel [1, 1]
                    let excel_start = if p_start < 0 {
                        // 负数索引：-1 代表最后一行数据
                        height as i32 + p_start + 1
                    } else {
                        p_start
                    };

                    let excel_end = if p_end < 0 {
                        // 负数索引：-1 代表最后一行数据
                        height as i32 + p_end + 1
                    } else {
                        p_end
                    };

                    // 检查开始行和结束行的合法性
                    if excel_start > excel_end {
                        // 开始行大于结束行，视为无效范围，不进行排除
                        return Ok(BooleanChunked::from_slice("mask".into(), &mask));
                    }

                    // 确保范围有效并在 Excel 行号范围内 (0 到 height)
                    let actual_start = excel_start.max(0).min(height as i32) as usize;
                    let actual_end = excel_end.max(0).min(height as i32) as usize;

                    // 排除指定范围的行
                    for i in actual_start..=actual_end {
                        if i <= height {
                            mask[i] = false;
                        }
                    }

                    Ok(BooleanChunked::from_slice("mask".into(), &mask))
                },
            }?; // 透過 ? 解開分支 Result 並進行錯誤轉換
            full_mask = full_mask & mask;
        }

        Ok(full_mask)
    }

    /// 评估列条件并生成匹配的列索引列表
    ///
    /// 专门用于处理列级别样式覆盖规则中的列条件判断，
    /// 返回满足所有条件的列索引集合。此函数仅处理与列位置相关的条件，
    /// 不涉及具体数据内容的判断。
    ///
    /// # 参数说明
    ///
    /// * `df`: 输入的 DataFrame，提供列结构信息
    /// * `conditions`: 列级别条件列表，仅包含位置相关条件
    ///
    /// # 返回值
    ///
    /// * `Ok(Vec<u16>)` - 匹配的列索引列表（去重并排序）
    ///   - 索引基于 0-based 系统
    ///   - u16 类型与 Excel 列地址系统对应
    /// * `Err(StyleFactoryError)` - 条件类型不匹配或其他错误
    fn evaluate_col_indices(&self, df: &DataFrame, conditions: &[StyleCondition]) -> Result<Vec<u16>, StyleFactoryError> {
        // 1. 嚴格檢查：禁止數據驅動型算子出現在列定義中
        let has_data_op = conditions.iter().any(|c| matches!(c, StyleCondition::ValueRange{..} | StyleCondition::Equal{..}));
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
                StyleCondition::Index { criteria, .. } => {
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

    /// 执行样式规则评估并生成单元格样式映射表
    ///
    /// 作为样式工厂的核心执行方法，此函数负责将预定义的样式规则
    /// 应用于输入的 DataFrame，生成完整的单元格级别样式映射。
    ///
    /// # 参数说明
    ///
    /// * `df`: 输入的 DataFrame，提供数据结构和维度信息
    ///
    /// # 返回值
    ///
    /// * `Ok(HashMap<(u32, u16), Arc<str>>)` - 单元格样式映射表
    ///   - Key: (行索引, 列索引) - 基于 0-based 系统
    ///   - Value: Arc 包装的样式名称字符串
    /// * `Err(Box<dyn std::error::Error>)` - 执行过程中的任何错误
    pub fn execute(&self, df: &DataFrame) -> Result<HashMap<(u32, u16), Arc<str>>, Box<dyn std::error::Error>> {
        // 初始化样式映射表：存储 (行,列) -> 样式名称 的映射
        let mut style_map = HashMap::new();
        // 获取 DataFrame 列数，用于表头和整行样式应用
        let width = df.width() as u16;

        // 遍历所有样式规则，按优先级顺序处理（后定义的规则覆盖前面的）
        for rule in &self.rules {
            // --- 【优化 1：使用 Arc 包装样式】 ---
            // 将基础样式名称包装为 Arc<str> 实现内存共享
            let base_style: Arc<str> = rule.apply.style.as_str().into();

            // 1. 获取行遮罩：确定当前规则作用于哪些行（包括表头行）
            // 返回 BooleanChunked，长度为 height + 1，索引直接对应 Excel 行号
            let row_mask = self.evaluate_conditions(df, &rule.row_conditions)?;

            // 2. 预处理 Overrides：预先计算所有覆盖规则的目标列
            let mut prepared_overrides = Vec::new();
            if let Some(ref overrides) = rule.apply.overrides {
                for ov in overrides {
                    // 直接评估列索引，让 evaluate_col_indices 的错误处理发挥作用
                    match self.evaluate_col_indices(df, &ov.col_conditions) {
                        Ok(target_cols) => {
                            let ov_style: Arc<str> = ov.style.as_str().into();
                            prepared_overrides.push((target_cols, ov_style));
                        }
                        Err(_) => {
                            // 包含数据驱动型条件，静默忽略
                            continue;
                        }
                    }
                }
            }

            for (r_idx, matched) in row_mask.into_iter().enumerate() {
                // 只处理匹配成功的行 (matched 为 Some(true))
                if let Some(true) = matched {
                    let r_phys = r_idx as u32;
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