use xlsx_writer::prelude::*;
use std::collections::HashMap;
use std::sync::Arc;

// 輔助函數：創建一個簡單的 DataFrame
fn create_mock_df() -> DataFrame {
    polars::df!(
        "姓名" => &["張三", "李四"],
        "得分" => &[90, 85]
    ).unwrap()
}

#[test]
fn test_empty_dataframe_skip() -> Result<(), Box<dyn std::error::Error>> {
    // 測試：空表應該被靜默跳過，不影響後續執行
    let empty_df = DataFrame::default();
    let normal_df = create_mock_df();

    let exporter = Workbook::new()?
        .insert(empty_df, Some("空表".into()), None)? // 應該跳過
        .insert(normal_df, Some("正常表".into()), None)?; // 應該成功

    // 驗證：最終隊列中應該只有 1 張表
    assert_eq!(exporter.sheets.len(), 1);
    assert_eq!(exporter.sheets[0].name, "正常表");
    Ok(())
}

#[test]
fn test_invalid_name_auto_fix() -> Result<(), Box<dyn std::error::Error>> {
    // 測試：包含非法字符的名稱應被自動修正為 "Sheet 1"
    let df = create_mock_df();
    let invalid_name = Some("工資/報表*2023".to_string());

    let exporter = Workbook::new()?
        .insert(df, invalid_name, None)?;

    // 驗證：名稱被重置為默認格式
    assert_eq!(exporter.sheets[0].name, "Sheet 1");
    Ok(())
}

#[test]
fn test_duplicate_name_error() -> Result<(), Box<dyn std::error::Error>> {
    // 測試：手動插入兩個同名表應返回 DuplicateName 錯誤
    let df1 = create_mock_df();
    let df2 = create_mock_df();

    let result = Workbook::new()?
        .insert(df1, Some("重複名".into()), None)?
        .insert(df2, Some("重複名".into()), None);

    // 驗證：返回了預期的重複名錯誤
    assert!(result.is_err());
    let err_msg = format!("{}", result.unwrap_err());
    assert!(err_msg.contains("重複的工作表名稱"));
    Ok(())
}

#[test]
fn test_unknown_style_error() -> Result<(), Box<dyn std::error::Error>> {
    // 測試：使用未定義的樣式標籤應報錯
    let df = create_mock_df();
    let mut style_map: HashMap<(u32, u16), Arc<str>> = HashMap::new();
    style_map.insert((1, 1), Arc::from("non_existent_style"));

    let result = Workbook::new()?
        .insert(df, None, Some(style_map));

    // 驗證：捕獲到未知樣式錯誤
    assert!(result.is_err());
    Ok(())
}

#[test]
fn test_style_coordinate_cleanup() -> Result<(), Box<dyn std::error::Error>> {
    // 測試：越界的樣式坐標應在 WorkSheet::new 中被自動剔除
    let df = create_mock_df(); // 2行2列
    let mut style_map = HashMap::new();
    style_map.insert((100, 100), Arc::from("header")); // 顯然越界

    let exporter = Workbook::new()?
        .insert(df, Some("坐標測試".into()), Some(style_map))?;

    // 驗證：style_map 最終變為 None（因為唯一的坐標被 retain 剔除了）
    assert!(exporter.sheets[0].style_map.is_none());
    Ok(())
}