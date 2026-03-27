use serde::Deserialize;
use std::collections::HashMap;
use rust_xlsxwriter::{Format, Color, FormatBorder, FormatAlign, FormatUnderline, FormatPattern};


/// 單個單元格樣式的詳細定義。
///
/// 該結構體作為中間層，將 JSON 的弱類型字符串轉換為 Excel 的強類型枚舉。
#[derive(Deserialize, Debug, Default, Clone)]
pub struct StyleDefinition {
    // --- 字體相關 ---
    pub font_name: Option<String>,
    pub font_size: Option<f64>,
    pub font_color: Option<String>, // 支持 "#RRGGBB"
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<String>,   // "single", "double" 等

    // --- 對齊相關 ---
    pub align: Option<String>,       // "left", "center", "right", "fill", "justify"
    pub valign: Option<String>,      // "top", "vcenter", "bottom", "vjustify"
    pub text_wrap: Option<bool>,

    // --- 邊框相關 (支持四周統一設置) ---
    pub border: Option<String>,      // "none", "thin", "medium", "dashed", "thick"
    pub border_color: Option<String>,

    // --- 背景與底紋 ---
    pub bg_color: Option<String>,
    pub fg_color: Option<String>,
    pub pattern: Option<u8>,         // 0..18 對應 Excel Pattern

    // --- 數據格式 ---
    pub num_format: Option<String>,  // 如 "yyyy-mm-dd" 或 "#,##0.00"
}

impl StyleDefinition {

    /// 將抽象的樣式定義轉換為 Excel 實體格式對象（Format）。
    ///
    /// 該方法透過逐一檢查可選字段，動態構建一個完整的 Excel 單元格格式。
    ///
    /// # 邏輯特點
    /// - **重新賦值模式**：針對底層庫的所有權轉移特性，採用 `format = format.method()` 確保狀態連續。
    /// - **默認回退**：未在 `StyleDefinition` 中指定的屬性將不會被調用，從而保留 Excel 的系統默認外觀。
    ///
    /// # 返回值
    /// 返回一個可直接用於 `Workbook` 寫入操作的 `rust_xlsxwriter::Format` 實例。
    pub fn to_format(&self) -> Format {
        // 始終從一個全新的 Format 開始
        let mut format = Format::new();

        // 1. 字體處理：注意賦值回 format 變量
        if let Some(ref name) = self.font_name { format = format.set_font_name(name); }
        if let Some(size) = self.font_size { format = format.set_font_size(size); }
        if let Some(ref color) = self.font_color { format = format.set_font_color(parse_color(color)); }

        // 對於 Boolean 標記，rust_xlsxwriter 的 API 通常也是返回新實例
        if let Some(true) = self.bold { format = format.set_bold(); }
        if let Some(true) = self.italic { format = format.set_italic(); }

        if let Some(ref u) = self.underline {
            let underline_type = match u.as_str() {
                "single" => FormatUnderline::Single,
                "double" => FormatUnderline::Double,
                _ => FormatUnderline::None,
            };
            format = format.set_underline(underline_type);
        }

        // 2. 對齊處理
        if let Some(ref h) = self.align {
            let h_align = match h.as_str() {
                "left" => FormatAlign::Left,
                "center" => FormatAlign::Center,
                "right" => FormatAlign::Right,
                "fill" => FormatAlign::Fill,
                "justify" => FormatAlign::Justify,
                _ => FormatAlign::General,
            };
            format = format.set_align(h_align);
        }

        if let Some(ref v) = self.valign {
            let v_align = match v.as_str() {
                "top" => FormatAlign::Top,
                "vcenter" | "center" => FormatAlign::VerticalCenter,
                "vjustify" | "justify" => FormatAlign::VerticalJustify,
                "vdistributed" | "distributed" => FormatAlign::VerticalDistributed,
                // 直接使用 Bottom 作為保底返回，涵蓋了 "bottom" 和所有非法輸入
                _ => FormatAlign::Bottom,
            };
            format = format.set_align(v_align);
        }

        if let Some(true) = self.text_wrap { format = format.set_text_wrap(); }

        // 3. 邊框處理
        if let Some(ref b) = self.border {
            let border_style = match b.as_str() {
                "thin" => FormatBorder::Thin,
                "medium" => FormatBorder::Medium,
                "dashed" => FormatBorder::Dashed,
                "thick" => FormatBorder::Thick,
                "double" => FormatBorder::Double,
                "hair" => FormatBorder::Hair,
                _ => FormatBorder::None,
            };
            format = format.set_border(border_style);
        }

        if let Some(ref color) = self.border_color {
            format = format.set_border_color(parse_color(color));
        }

        // 4. 背景與模式
        if let Some(ref bg) = self.bg_color {
            format = format.set_background_color(parse_color(bg));
        }
        if let Some(ref fg) = self.fg_color {
            format = format.set_foreground_color(parse_color(fg));
        }

        if let Some(p) = self.pattern {
            let pattern_type = match p {
                0 => FormatPattern::None,
                1 => FormatPattern::Solid, // 最常用的實色填充
                2 => FormatPattern::MediumGray,
                3 => FormatPattern::DarkGray,
                4 => FormatPattern::LightGray,
                5 => FormatPattern::DarkHorizontal,
                6 => FormatPattern::DarkVertical,
                // ... 如果有需要可以繼續添加，或者直接默認 Solid
                _ => FormatPattern::Solid,
            };
            format = format.set_pattern(pattern_type);
        }

        // 5. 數字格式
        if let Some(ref f) = self.num_format {
            format = format.set_num_format(f);
        }

        format
    }

}

/// 樣式資源庫：存儲從 JSON 加載的所有樣式模板。
#[derive(Deserialize, Debug, Default)]
pub struct StyleLibrary {
    /// 樣式映射表。
    ///
    /// 鍵 (Key): 樣式的唯一標識符（Label），供 JSON 規則或代碼引用。
    /// 值 (Value): 對應的樣式配置細節（StyleDefinition）。
    ///
    /// 該字段支持透過 Serde 直接從 JSON 對象中的 "styles" 鍵位反序列化。
    pub styles: HashMap<String, StyleDefinition>,
}

impl StyleLibrary {

    /// 創建一個空的樣式資源庫實例。
    ///
    /// 用於手動構建樣式定義或作為動態插入的基礎容器。
    pub fn new() -> Self {
        Self {
            styles: HashMap::new(),
        }
    }

    /// 從 JSON 對象片段解析並填充樣式庫。
    ///
    /// # 參數
    /// * `value` - 指向樣式配置片段的引用（預期為 JSON Object 結構）。
    ///
    /// # 錯誤處理
    /// 若 `value` 的結構無法匹配 `HashMap<String, StyleDefinition>`，將拋出反序列化異常。
    pub fn from_json(value: &serde_json::Value) -> Result<Self, Box<dyn std::error::Error>> {
        // 將 JSON 片段轉換為強類型的映射表。使用 clone 確保不影響原始 Value 的所有權。
        // 如果傳入的不是 Object，serde_json 會在此處拋出錯誤，這符合「輸入無效則報錯」的預期
        let definitions: HashMap<String, StyleDefinition> = serde_json::from_value(value.clone())?;

        Ok(Self { styles: definitions })
    }

    /// 手動插入或更新單個樣式定義。
    ///
    /// # 參數
    /// * `name` - 樣式標籤。
    /// * `definition` - 樣式的詳細屬性。
    ///
    /// # 返回值
    /// 返回修改後的實例，支持鏈式調用。
    pub fn insert(mut self, name: &str, definition: StyleDefinition) -> Self {
        self.styles.insert(name.to_string(), definition);
        self
    }

    /// 從 JSON 片段批量插入或更新多個樣式定義。
    ///
    /// # 參數
    /// * `value` - 指向 JSON Object 的引用，其鍵應為樣式名。
    ///
    /// # 錯誤處理
    /// 若傳入的 JSON 片段與映射表結構不符（例如非對象格式），則返回反序列化錯誤。
    pub fn insert_from_json(self, value: &serde_json::Value) -> Result<Self, Box<dyn std::error::Error>> {
        // 1. 檢查是否為有效的 JSON Object
        if !value.is_object() {
            // 返回一個空的庫實例，並可選擇性打印警告
            eprintln!("Warning: StyleLibrary input is not a valid object, using empty library.");
            return Ok(Self::new());
        }

        // 2. 只有是 Object 時才執行解析
        let definitions: HashMap<String, StyleDefinition> = serde_json::from_value(value.clone())?;
        Ok(Self { styles: definitions })
    }

    /// 將資源庫中所有的抽象定義轉換為物理樣式對象池。
    ///
    /// 該方法會遍歷 `styles` 映射表，調用每個定義的轉換邏輯，並產出一個可供 Excel 引擎直接使用的映射。
    ///
    /// # 返回值
    /// 返回一個 `HashMap<String, Format>`，其鍵與資源庫中的標籤一致，值為預配置的 Excel 格式。
    pub fn build_formats(&self) -> HashMap<String, Format> {
        self.styles
            .iter()
            .map(|(name, def)| (name.clone(), def.to_format()))
            .collect()
    }

}

/// 將 HTML/CSS 風格的十六進位顏色字串解析為 Excel 顏色對象。
///
/// 該輔助函數執行預處理與進制轉換，將字符串標識轉化為 24 位 RGB 數值。
///
/// # 解析流程
/// 1. **清洗**：移除首尾空格。
/// 2. **去綴**：過濾掉起始的 '#' 字符。
/// 3. **轉換**：以 16 進位解析剩餘字符為 `u32` 整數。
/// 4. **封裝**：若成功則構造 `Color::RGB`，若失敗則回退至系統默認色。
fn parse_color(hex: &str) -> Color {
    // 預處理：支持 "#RRGGBB"、"RRGGBB" 以及帶空格的輸入
    let hex_clean = hex.trim().trim_start_matches('#');

    // 將 16 進位字串轉換為 24-bit 顏色整數 (0xRRGGBB)
    if let Ok(rgb) = u32::from_str_radix(hex_clean, 16) {
        Color::RGB(rgb)
    } else {
        // 容錯邏輯：遇到非法字元時不崩潰，保持 Excel 默認樣式
        Color::Default
    }
}