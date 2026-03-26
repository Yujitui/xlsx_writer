use std::collections::HashMap;
use std::error::Error;
use polars::prelude::*;
use rust_xlsxwriter::*;
use crate::worksheet::WorkSheet;
use crate::error::XlsxError;
use std::fmt;

/// Excel 導出管理器（核心容器）。
///
/// 該結構體負責維護全局樣式配置以及待導出的工作表隊列。
/// 通過生命週期管理，它將多個 DataFrame 任務整合到單個物理 `.xlsx` 文件中。
pub struct Workbook {

    /// 全局樣式集（樣式池）。
    ///
    /// 鍵（Key）為樣式標籤（如 "header", "money"），值為預配置的 `rust_xlsxwriter::Format` 對象。
    /// 在寫入單元格時，通過標籤引用此池中的樣式，可大幅減少內存開銷並確保報表風格統一。
    pub styles: HashMap<String, Format>,
    /// 待執行的工作表導出任務隊列。
    ///
    /// 每個 [`WorkSheet`] 實例代表 Excel 中的一個 Sheet。
    /// 導出時將按照此 Vec 中的順序進行物理寫入。
    pub sheets: Vec<WorkSheet>,

}

impl Workbook {

    /// 創建一個新的 `Workbook` 實例並初始化基礎樣式池。
    ///
    /// 該函數會根據編譯目標平台自動選擇最適宜的中文字體，並預置「標準數據」與「灰色表頭」兩種樣式。
    ///
    /// # 邏輯流程
    /// 1. **字體適配**：Windows 使用「微軟雅黑」，macOS 使用「苹方」，其餘平台回退至通用無襯線字體。
    /// 2. **基礎樣式 (`default`)**：定義 11 號字、細黑邊框、水平垂直居中的通用樣式。
    /// 3. **衍生樣式 (`header`)**：基於基礎樣式，增加加粗效果與 `#BFBFBF` 淺灰色背景。
    ///
    /// # 錯誤 (Returns)
    /// * `Result<Self, Box<dyn Error>>` - 理論上在內存充足時不會報錯，返回 `Result` 是為了保持與後續 IO 操作的接口一致性。
    pub fn new() -> Result<Self, Box<dyn Error>> {
        let mut styles = HashMap::new();

        // 1. 跨平台字體自適應適配：
        // 利用編譯時宏 cfg! 確定目標系統。Excel 文件不嵌入字體，僅存儲名稱，
        // 故需確保指定的名稱在對應系統中是可被識別的標準名稱。
        let font_name = if cfg!(target_os = "windows") {
            "Microsoft YaHei" // Windows: 微软雅黑
        } else if cfg!(target_os = "macos") {
            "PingFang SC"      // macOS: 苹方
        } else {
            "sans-serif"       // Linux 或其他: 通用無襯線字體
        };

        // 2. 定義「標準樣式」：
        // 這是報表中絕大多數單元格的底層模板，包含了邊框、字體和對齊方式。
        let standard_fmt = Format::new()
            .set_font_name(font_name) // 設置微軟雅黑
            .set_font_size(11)                 // 設置常用字號
            .set_border(FormatBorder::Thin)   // 設置細邊框（四周）
            .set_align(FormatAlign::Center)   // 水平居中
            .set_align(FormatAlign::VerticalCenter); // 垂直居中

        // 3. 快速派生「表頭樣式」：
        // 通過 clone() 繼承 standard_fmt 的屬性，僅對差異項（加粗、背景）進行覆蓋。
        // 這種鏈式調用確保了表頭與數據行在字體、邊框寬度上完全對齊。
        let header_fmt = standard_fmt.clone()
            .set_bold()
            .set_background_color(Color::RGB(0xBFBFBF));

        // 4. 將初始化樣式存入緩存池：
        // "default" 用於普通單元格保底，"header" 用於標題行。
        styles.insert("default".to_string(), standard_fmt);
        styles.insert("header".to_string(), header_fmt);

        Ok(Self {
            styles,
            sheets: vec![],
        })
    }

    /// 向樣式池中插入或更新一個自定義樣式。
    ///
    /// 該方法支持鏈式調用，允許連續配置多個樣式。
    ///
    /// # 參數
    /// * `name` - 樣式的唯一標識符（標籤）。建議使用具備語義的名稱（如 "money", "warning"）。
    /// * `format` - 預配置的 `rust_xlsxwriter::Format` 對象。
    ///
    /// # 行為說明
    /// - **覆蓋邏輯**：如果傳入的 `name` 在樣式池中已存在，則會直接使用新樣式覆蓋舊樣式，並在後續導出中生效。
    /// - **所有權轉移**：接收 `mut self` 並返回 `Self`，方便在初始化時流暢配置。
    pub fn set_style(mut self, name: &str, format: Format) -> Self {
        // 將字符串切片轉換為 String 並存入 HashMap。
        // HashMap 的 insert 特性確保了同名鍵值的自動更新。
        self.styles.insert(name.to_string(), format);

        // 返回修改後的對象，支持鏈式調用：exporter.set_style(...).set_style(...)
        self
    }

    /// 將一個 DataFrame 數據源包裝為工作表任務並插入導出隊列。
    ///
    /// 該方法具備自動校驗與修復功能，支持鏈式調用。
    ///
    /// # 參數
    /// * `df` - 數據源。利用 Polars 的引用計數特性，內部會進行必要的淺拷貝。
    /// * `name` - 可選的工作表名稱。若為 `None` 或無效，將自動生成「Sheet N」格式的名稱。
    /// * `style_map` - 可選的單元格樣式坐標映射。
    ///
    /// # 業務邏輯與容錯
    /// 1. **空表過濾**：若 `df` 為空，該方法會**靜默跳過**插入並返回 `Ok(self)`，不會中斷後續操作。
    /// 2. **名稱修復**：若傳入的 `name` 包含非法字符或超長，將自動回退為默認的「Sheet N」名稱。
    /// 3. **衝突檢查**：若工作表名稱與隊列中已有的名稱重複，則視為邏輯錯誤並返回 `Err`。
    /// 4. **樣式校驗**：確保 `style_map` 引用的是已在 `Workbook.styles` 池中定義的樣式。
    ///
    /// # 錯誤 (Returns)
    /// * [`XlsxError::DuplicateName`] - 名稱與現有工作表衝突。
    /// * [`XlsxError::UnknownStyle`] - 引用了未定義的樣式標籤。
    /// * `Box<dyn Error>` - 其他未預期的構建錯誤。
    pub fn insert(mut self, df: DataFrame, name: Option<String>, style_map: Option<HashMap<(u32, u16), String>>) -> Result<Self, Box<dyn Error>> {
        // 1. 定義輔助閉包：封裝默認命名邏輯，確保命名的一致性與唯一性起點
        let get_default_name = |sheets_len: usize| format!("Sheet {}", sheets_len + 1);

        // 2. 初步確定名稱：優先使用用戶提供，否則生成默認名
        let final_name = name.unwrap_or_else(|| get_default_name(self.sheets.len()));

        // 3. 嘗試構建 WorkSheet 任務：
        // 這裡使用 match 進行細分錯誤處理。注意 df.clone() 是淺拷貝，開銷極低。這裡 clone() 是必須的，因為如果 InvalidName 發生，我們需要原始數據進行第二次嘗試。
        let task = match WorkSheet::new(df.clone(), final_name.clone(), style_map.clone()) {
            Ok(t) => t,
            // 規則 A (靜默跳過)：空表不具備導出意義，直接返回原始對象，不存入隊列
            Err(XlsxError::EmptyDataFrame) => return Ok(self),
            // 規則 B (自動修復)：名稱非法時，放棄用戶名稱，改用系統預設名稱重試
            // 此時已知 df 不為空，第二次 new 操作是安全的
            Err(XlsxError::InvalidName(_)) => {
                let fallback_name = get_default_name(self.sheets.len());
                // 再次調用 new，此時使用安全名稱（已知 df 不為空，所以這次一定會 Ok）
                WorkSheet::new(df, fallback_name, style_map)?
            }
            // 其他嚴重錯誤：直接包裝並向上拋出
            Err(e) => return Err(Box::new(e)),
        };

        // 4. 名稱重複檢查：
        // Excel 不允許同名工作表。這是全局級別的衝突，必須由 Workbook 攔截。
        if self.sheets.iter().any(|s| s.name == task.name) {
            return Err(Box::new(XlsxError::DuplicateName(task.name)));
        }

        // 5. 樣式名存在性檢查：
        // 確保數據在寫入時能找到對應的格式定義，防止 save 時出現懸空引用。
        if let Some(ref map) = task.style_map {
            for style_name in map.values() {
                if !self.styles.contains_key(style_name) {
                    return Err(Box::new(XlsxError::UnknownStyle(style_name.clone())));
                }
            }
        }

        // 通過所有校驗，將任務存入隊列
        self.sheets.push(task);
        Ok(self)
    }

    /// 執行物理導出，將所有工作表寫入指定路徑。
    ///
    /// # 參數
    /// * `path` - 目標文件路徑（如 "report.xlsx"）。
    ///
    /// # 邏輯流程
    /// 1. **初始化引擎**：創建 `rust_xlsxwriter` 實例。
    /// 2. **樣式準備**：獲取全局 `default` 樣式，並準備一個 `fallback` 格式以應對借用檢查。
    /// 3. **雙重循環遍歷**：
    ///     - **外層**：遍歷各個 `WorkSheet` 任務。
    ///     - **中層**：遍歷 DataFrame 的每一列（Column）。
    ///     - **內層**：遍歷每一行數據，將 `AnyValue` 映射為 Excel 類型。
    /// 4. **樣式應用**：
    ///     - **優先級**：`style_map` 指定標籤 > 全局 `default` > 系統 `fallback`。
    ///     - 確保表頭（Row 0）與數據行（Row 1..N）均能正確應用格式。
    /// 5. **自動列寬**：寫入完畢後自動調整寬度以適應字體。
    ///
    /// # 錯誤 (Returns)
    /// * `Box<dyn Error>` - 包含 Polars 數據讀取錯誤、Excel 引擎寫入錯誤或 IO 權限錯誤。
    pub fn save(&self, path: &str) -> Result<(), Box<dyn Error>> {
        let mut workbook = rust_xlsxwriter::Workbook::new();

        // --- 1. 樣式池安全預取： ---
        // fallback_fmt 用於解決 Rust 借用檢查器對臨時引用的限制。
        // default_fmt 是你在 new() 中定義的全局 UI 規範（如微軟雅黑+邊框）。
        let fallback_fmt = Format::default();
        let default_fmt = self.styles.get("default");

        for sheet in &self.sheets {
            let worksheet = workbook.add_worksheet();
            worksheet.set_name(sheet.name.as_str())?;

            // 獲取 Polars 列引用
            let columns = sheet.df.columns();

            for (col_idx, column) in columns.iter().enumerate() {
                let c = col_idx as u16;

                // --- 2. 處理表頭 (Excel Row 0) ---
                // 邏輯升級：
                // 1. 優先查找 style_map 中的定義。
                // 2. 如果 style_map 是 None，則自動應用預設的 "header" 樣式。
                // 3. 如果以上皆否，則使用全局 "default"。
                // 4. 最後使用系統 "fallback"。

                let header_cell_fmt = sheet.style_map.as_ref()
                    .and_then(|m| m.get(&(0, c)))         // 嘗試從地圖找
                    .and_then(|name| self.styles.get(name))
                    .or_else(|| {                        // 如果地圖沒定義或地圖不存在
                        if sheet.style_map.is_none() {
                            self.styles.get("header")    // 自動應用預設 header
                        } else {
                            None
                        }
                    })
                    .or(default_fmt)                     // 全局保底
                    .unwrap_or(&fallback_fmt);           // 系統保底
                worksheet.write_with_format(0, c, column.name().as_str(), header_cell_fmt)?;

                // --- 3. 處理數據行 (Excel Row 1..N) ---
                for row_idx in 0..sheet.df.height() {
                    let val = column.get(row_idx)?; // 获取 AnyValue
                    let r = (row_idx + 1) as u32;   // Excel 行索引（跳过表头）

                    // 獲取該單元格專屬樣式或全局保底樣式
                    let cell_fmt = sheet.style_map.as_ref()
                        .and_then(|m| m.get(&(r, c)))
                        .and_then(|name| self.styles.get(name))
                        .or(default_fmt)
                        .unwrap_or(&fallback_fmt); // 同理

                    // --- 4. 類型分派：將 Polars 豐富的數值類型統化為 Excel 的 f64。 ---
                    // 使用 _with_format 系列方法確保樣式（邊框、字體）被正確應用。
                    match val {
                        // 使用 _with_format 系列方法，傳入 cell_fmt (&Format)
                        AnyValue::Int8(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Int16(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Int32(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Int64(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::UInt32(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Float32(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Float64(v) => worksheet.write_number_with_format(r, c, v, cell_fmt)?,
                        AnyValue::String(s) => worksheet.write_string_with_format(r, c, s, cell_fmt)?,
                        AnyValue::Boolean(v) => worksheet.write_boolean_with_format(r, c, v, cell_fmt)?,
                        AnyValue::Null => {
                            // Null 值也寫入 blank 以保持單元格邊框一致
                            worksheet.write_blank(r, c, cell_fmt)?
                        },
                        _ => {
                            // 處理日期等類型，先轉為字符串
                            let s = format!("{}", val);
                            worksheet.write_string_with_format(r, c, &s, cell_fmt)?
                        },
                    };
                }
            }

            // --- 5. 自動列寬適配：---
            // 必須在數據寫入完成後調用，確保 Excel 引擎能計算出最長內容的佔位。
            worksheet.autofit();
        }

        workbook.save(path)?;
        Ok(())
    }

}

impl fmt::Debug for Workbook {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Workbook")
            .field("sheets", &self.sheets) // 只打印實現了 Debug 的 sheets
            .field("styles_count", &self.styles.len()) // 打印樣式數量作為替代
            .finish()
    }
}