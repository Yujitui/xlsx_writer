#[cfg(test)]
mod tests {
    use xlsx_writer::prelude::*;
    use polars::prelude::*;

    #[test]
    fn test_style_factory_matrix_logic() -> Result<(), Box<dyn std::error::Error>> {
        // 1. 準備測試數據 (DataFrame)
        // 索引 0: 張三 (研發部, 12000) -> 物理行 1
        // 索引 1: 李四 (財務部, 4500)  -> 物理行 2
        // 索引 2: 王五 (研發部, 15000) -> 物理行 3
        let df = df!(
            "姓名" => ["张三", "李四", "王五"],
            "部门" => ["研发部", "财务部", "研发部"],
            "实发工资" => [12000.0, 4500.0, 15000.0]
        )?;

        // 2. 定義複雜的 JSON 規則
        let json_config = r#"[
                {
                    "row_conditions": [{ "type": "index", "criteria": [0] }],
                    "apply": {
                        "style": "header_bg",
                        "overrides": [
                            {
                                "style": "header_money_highlight",
                                "col_conditions": [{ "type": "match", "targets": ["实发工资"], "criteria": [] }]
                            }
                        ]
                    }
                },
                {
                    "row_conditions": [
                        { "type": "exclude_rows", "criteria": [1, 1] },
                        { "type": "value_range", "targets": ["实发工资"], "criteria": ">=10000" },
                        { "type": "match", "targets": ["部门"], "criteria": ["研发部"] }
                    ],
                    "apply": {
                        "style": "high_salary_row",
                        "overrides": [
                            {
                                "style": "money_red_bold",
                                "col_conditions": [{ "type": "match", "targets": ["实发工资"], "criteria": [] }]
                            }
                        ]
                    }
                }
        ]"#;

        // 3. 執行工廠引擎
        let factory = StyleFactory::from_json_str(json_config)?;
        let style_map = factory.execute(&df)?;

        // --- 驗證點 1：物理第 0 行 (標題行) ---
        // 基礎樣式應為 "header_bg"
        assert_eq!(style_map.get(&(0, 0)).map(|s| s.as_ref()), Some("header_bg"));
        // 「實發工資」列 (索引 2) 應被 Override 為 "header_money_highlight"
        assert_eq!(style_map.get(&(0, 2)).map(|s| s.as_ref()), Some("header_money_highlight"));

        // --- 驗證點 2：物理第 1 行 (張三, DF索引 0) ---
        // 雖然張三符合高薪+研發部，但他被 exclude_rows:  (物理第1行) 排除掉了。
        // 所以他在 Map 中不應該有任何由「高薪規則」產生的樣式。
        assert!(style_map.get(&(1, 0)).is_none(), "張三應該被物理行排除邏輯過濾");

        // --- 驗證點 3：物理第 2 行 (李四, DF索引 1) ---
        // 李四不符合高薪條件，應為 None
        assert!(style_map.get(&(2, 0)).is_none());

        // --- 驗證點 4：物理第 3 行 (王五, DF索引 2) ---
        // 王五符合條件 (不是第1行, 研發部, 15000 >= 10000)
        // 基礎樣式
        assert_eq!(style_map.get(&(3, 0)).map(|s| s.as_ref()), Some("high_salary_row"));
        // 局部 Override
        assert_eq!(style_map.get(&(3, 2)).map(|s| s.as_ref()), Some("money_red_bold"));

        println!("樣式引擎測試通過！");
        Ok(())
    }
}