My Xlsx Exporter
一個基於 Polars 和 rust_xlsxwriter 構建的 Rust 高性能 Excel 導出工具。
✨ 特性
高性能：利用 Polars 的內存管理（Arc 引用計數）實現 DataFrame 的零成本轉移。
跨平台字體自適應：自動為 Windows（微軟雅黑）和 macOS（苹方）匹配最佳中文字體。
強大容錯：自動修復非法工作表名稱（如包含 :/ \* [ ] 等字符），自動跳過空數據表。
樣式解耦：預設「標準」與「灰色表頭」樣式池，支持通過坐標映射（StyleMap）實現精準單元格控制。
類型安全：完善的錯誤枚舉體系，支持數據類型、重複名及樣式名校驗。
自動列寬：一鍵 autofit，確保內容不被遮擋。
📦 安裝
在你的項目 Cargo.toml 中引用此 GitHub 倉庫：
toml
[dependencies]
# 引用特定版本（推薦）
my_xlsx_exporter = { git = "https://github.com", tag = "v0.1.0" }

# 或者引用最新代碼
# my_xlsx_exporter = { git = "https://github.com" }
请谨慎使用此类代码。

🚀 快速上手
使用 prelude 模塊可以快速獲取所有核心類型。
rust
use my_xlsx_exporter::prelude::*;
use std::collections::HashMap;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 1. 準備數據 (Polars DataFrame)
    let df = polars::df!(
        "姓名" => &["張三", "李四", "王五"],
        "薪資" => &[15000.5, 18000.0, 22000.75],
        "是否代繳" => &[true, false, true]
    )?;

    // 2. (可選) 定義特殊單元格樣式
    // 坐標為 (行, 列)，(0, 0) 代表第一列的表頭
    let mut style_map = HashMap::new();
    style_map.insert((1, 1), "money".to_string()); // 為張三的薪資應用自定義樣式

    // 3. 構建並執行導出
    Workbook::new()?
        // 插入工作表：數據, 名稱(None則自動命名), 樣式表
        .insert(df, Some("2023工資表".into()), Some(style_map))?
        // 支持鏈式調用插入多個 Sheet
        .insert(polars::df!("A" => &[1, 2])?, None, None)? 
        // 執行保存
        .save("財務報表_2023.xlsx")?;

    println!("✨ Excel 導出成功！");
    Ok(())
}
请谨慎使用此类代码。

🛠️ 預設樣式池
本工具默認提供以下樣式標籤，可直接在 style_map 中使用：
標籤名	說明
default	標準字體（微軟雅黑/苹方）、細黑邊框、居中對齊。
header	在 default 基礎上增加加粗及淺灰色背景 (#BFBFBF)。
你也可以通過 set_style 方法擴展自己的樣式池。
⚠️ 注意事項
Sheet 名稱限制：Excel 限制工作表名稱不得超過 31 個字符。若名稱非法或重複，本工具會自動回退至 Sheet N 格式以確保導出成功。
依賴版本：請確保你的項目使用的 polars 版本與本庫一致（目前為 0.38），以避免類型不匹配錯誤。
