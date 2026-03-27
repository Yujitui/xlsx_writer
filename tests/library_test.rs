#[cfg(test)]
mod tests {
    use xlsx_writer::prelude::*;

    #[test]
    fn test_style_library_full_flow() -> Result<(), Box<dyn std::error::Error>> {
        // 1. 模擬從外部加載的樣式 JSON 片段
        let json_data = serde_json::json!({
            "header": {
                "font_name": "Arial",
                "bold": true,
                "bg_color": "#BFBFBF",
                "align": "center",
                "valign": "vcenter",
                "border": "thin"
            },
            "money": {
                "font_color": "#FF0000",
                "num_format": "#,##0.00"
            }
        });

        // 2. 測試 from_json：解析 JSON 片段
        let mut library = StyleLibrary::from_json(&json_data)?;

        // 驗證解析數量
        assert_eq!(library.styles.len(), 2);
        assert!(library.styles.contains_key("header"));

        // 3. 測試 insert：手動動態增加一個樣式
        let custom_def = StyleDefinition {
            font_size: Some(14.0),
            italic: Some(true),
            ..Default::default()
        };
        library = library.insert("custom", custom_def);
        assert_eq!(library.styles.len(), 3);

        // 4. 測試 build_formats：物質化為 Excel Format 對象
        let formats = library.build_formats();

        // 5. 深度驗證轉換後的物理屬性

        // 驗證 "header" 樣式
        let _header_fmt = formats.get("header").expect("應該存在 header 樣式");
        // 注意：rust_xlsxwriter 的 Format 屬性通常不對外暴露 getter，
        // 但我們可以通過 Workbook 嘗試寫入來驗證，或者在此處確保轉換過程不崩潰。
        // 以下斷言基於我們之前的邏輯實現：
        assert!(library.styles.get("header").unwrap().bold.unwrap());

        // 驗證顏色解析
        let money_def = library.styles.get("money").unwrap();
        assert_eq!(money_def.font_color.as_ref().unwrap(), "#FF0000");

        // 6. 測試 insert_from_json：批量覆蓋/更新
        let patch_json = serde_json::json!({
            "money": { "font_color": "#00FF00" }, // 覆蓋：紅字變綠字
            "new_style": { "italic": true }       // 新增
        });
        library = library.insert_from_json(&patch_json)?;

        assert_eq!(library.styles.len(), 4);
        // 驗證覆蓋結果
        assert_eq!(library.styles.get("money").unwrap().font_color.as_ref().unwrap(), "#00FF00");

        println!("StyleLibrary 矩陣資源測試全部通過！");
        Ok(())
    }
}