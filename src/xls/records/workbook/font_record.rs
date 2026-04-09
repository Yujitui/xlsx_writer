use super::encode_biff_string_v1;
use super::BiffRecord;

/// 字体结构体，定义字体的各种属性
///
/// ## 作用
///
/// Font结构体定义了Excel中使用的字体属性，包括字体名称、大小、颜色、粗体、斜体、下划线等。
/// 这些属性通过FontRecord记录存储在工作簿中，每个工作簿最多可以有4个字体（Excel BIFF8限制）。
///
/// ## 参数说明
///
/// - `height`: 字体高度（单位：1/20点）
/// - `options`: 字体选项标志位
///   - bit 0 (0x0001): 粗体
///   - bit 1 (0x0002): 斜体
///   - bit 2 (0x0004): 下划线
///   - bit 3 (0x0008): 删除线
/// - `colour_index`: 颜色索引（0x7FFF为自动颜色）
/// - `weight`: 字重（0x0190=常规, 0x02BC=粗体）
/// - `escapement`: 上标/下标（0x0000=普通, 0x0001=上标, 0x0002=下标）
/// - `underline`: 下划线样式
///   - 0x00: 无下划线
///   - 0x01: 单下划线
///   - 0x02: 双下划线
///   - 0x21: 会计单下划线
///   - 0x22: 会计双下划线
/// - `family`: 字体族（0x00=不需要, 0x01=衬线, 0x02=无衬线, 0x03=手写体, 0x04=等宽）
/// - `charset`: 字符集（0x01=Latin1, 0x00=ANSI）
/// - `name`: 字体名称
#[derive(Debug, Clone)]
pub struct Font {
    pub height: u16,
    pub options: u16,
    pub colour_index: u16,
    pub weight: u16,
    pub escapement: u16,
    pub underline: u8,
    pub family: u8,
    pub charset: u8,
    pub name: String,
}

impl Font {
    pub fn new(name: &str) -> Self {
        Font {
            height: 0x00C8,
            options: 0x0000,
            colour_index: 0x7FFF,
            weight: 0x0190,
            escapement: 0x0000,
            underline: 0x00,
            family: 0x00,
            charset: 0x01,
            name: name.to_string(),
        }
    }

    pub fn with_height(mut self, height: u16) -> Self {
        self.height = height;
        self
    }

    pub fn with_bold(mut self) -> Self {
        self.options |= 0x0001;
        self.weight = 0x02BC;
        self
    }

    pub fn with_italic(mut self) -> Self {
        self.options |= 0x0002;
        self
    }

    pub fn with_underline(mut self, underline: u8) -> Self {
        self.options |= 0x0004;
        self.underline = underline;
        self
    }

    pub fn with_struck_out(mut self) -> Self {
        self.options |= 0x0008;
        self
    }
}

impl Default for Font {
    fn default() -> Self {
        Font::new("Arial")
    }
}

/// FontRecord 记录
///
/// 作用：存储字体定义信息
///
/// FontRecord是Excel BIFF格式中的字体记录（ID: 0x0031），用于在工作簿中定义字体样式。
/// 每个FontRecord包含一个Font结构体，定义了字体的完整属性。
pub struct FontRecord {
    font: Font,
}

impl FontRecord {
    pub fn new(font: Font) -> Self {
        FontRecord { font }
    }
}

impl BiffRecord for FontRecord {
    fn id(&self) -> u16 {
        0x0031
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::new();

        // 2 bytes: height
        buf.extend_from_slice(&self.font.height.to_le_bytes());

        // 2 bytes: options
        buf.extend_from_slice(&self.font.options.to_le_bytes());

        // 2 bytes: colour index
        buf.extend_from_slice(&self.font.colour_index.to_le_bytes());

        // 2 bytes: weight
        buf.extend_from_slice(&self.font.weight.to_le_bytes());

        // 2 bytes: escapement
        buf.extend_from_slice(&self.font.escapement.to_le_bytes());

        // 1 byte: underline
        buf.push(self.font.underline);

        // 1 byte: family
        buf.push(self.font.family);

        // 1 byte: charset
        buf.push(self.font.charset);

        // 1 byte: reserved
        buf.push(0x00);

        // Font name using upack1 logic
        let encoded_name = encode_biff_string_v1(&self.font.name);
        buf.extend_from_slice(&encoded_name);

        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_font_record_id() {
        let record = FontRecord::new(Font::new("Arial"));
        assert_eq!(record.id(), 0x0031);
    }

    #[test]
    fn test_font_record_basic() {
        let record = FontRecord::new(Font::new("Arial"));
        let data = record.data();

        // Check height (2 bytes) - xlwt default is 200 (0x00C8)
        assert_eq!(&data[0..2], &0x00C8u16.to_le_bytes());

        // Check options (2 bytes)
        assert_eq!(&data[2..4], &0x0000u16.to_le_bytes());

        // Check colour index (2 bytes) - xlwt default is 0x7FFF
        assert_eq!(&data[4..6], &0x7FFFu16.to_le_bytes());

        // Check name starts at offset 14
        // name length at byte 14, flag at byte 15, name starts at byte 16
        assert_eq!(data[14], 5); // "Arial" = 5 chars
        assert_eq!(data[15], 0x00); // compressed ASCII
    }

    #[test]
    fn test_font_record_unicode() {
        let record = FontRecord::new(Font::new("微软雅黑"));
        let data = record.data();

        // Unicode name uses flag 0x01
        assert_eq!(data[15], 0x01);
    }
}
