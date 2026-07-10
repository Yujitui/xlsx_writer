//! 从 Excel 2026 参考文件（工作簿1.xls）中提取的默认 BIFF 记录数据
//! 
//! 这些数据用于替换生成文件中的默认值，以精确匹配 Excel 的输出格式。
//! 提取方式：解析参考文件的 Workbook CFB 流，逐记录提取二进制载荷。

use crate::xls_records::workbook::xf_record::{
    Alignment, Borders, Pattern, XF, XFRecord, XFType,
};
use crate::xls_records::BiffRecord;
use crate::xls_records::NumberFormatRecord;

/// 将十六进制字符串解码为字节向量
fn hex_to_bytes(hex: &str) -> Vec<u8> {
    (0..hex.len()).step_by(2)
        .map(|i| u8::from_str_radix(&hex[i..i+2], 16).unwrap())
        .collect()
}

/// 以十六进制字符串存储载荷的 BIFF 记录
pub struct HexBiffRecord {
    pub id: u16,
    pub hex: &'static str,
}

impl HexBiffRecord {
    pub const fn new(id: u16, hex: &'static str) -> Self {
        Self { id, hex }
    }
    pub fn serialize(&self) -> Vec<u8> {
        let data = hex_to_bytes(self.hex);
        let len = data.len() as u16;
        let mut buf = Vec::with_capacity(4 + data.len());
        buf.extend_from_slice(&self.id.to_le_bytes());
        buf.extend_from_slice(&len.to_le_bytes());
        buf.extend_from_slice(&data);
        buf
    }
}

pub fn serialize_default_fonts() -> Vec<u8> {
    use super::font_record::{Font, FontRecord};
    use crate::xls_records::BiffRecord;

    fn dx(height: u16, colour: u16, weight: u16, options: u16, name: &str) -> FontRecord {
        FontRecord::new(Font {
            height,
            options,
            colour_index: colour,
            weight,
            escapement: 0,
            underline: 0,
            family: 0,
            charset: 0x86,
            name: name.to_string(),
        })
    }

    let mut buf = Vec::new();
    let fonts = [
        // [0-4] 5× 12pt 等线 colour=0x08
        dx(0x00F0, 0x0008, 0x0190, 0x0000, "等线"),
        dx(0x00F0, 0x0008, 0x0190, 0x0000, "等线"),
        dx(0x00F0, 0x0008, 0x0190, 0x0000, "等线"),
        dx(0x00F0, 0x0008, 0x0190, 0x0000, "等线"),
        dx(0x00F0, 0x0008, 0x0190, 0x0000, "等线"),
        // [5] 9pt 等线 auto-colour
        dx(0x00B4, 0x7FFF, 0x0190, 0x0000, "等线"),
        // [6] 12pt 等线 colour=0x08 (same as 0-4)
        dx(0x00F0, 0x0008, 0x0190, 0x0000, "等线"),
        // [7] 18pt 等线 Light colour=0x38
        dx(0x0168, 0x0038, 0x0190, 0x0000, "等线 Light"),
        // [8] 15pt bold colour=0x38
        dx(0x012C, 0x0038, 0x02BC, 0x0001, "等线"),
        // [9] 13pt bold colour=0x38
        dx(0x0104, 0x0038, 0x02BC, 0x0001, "等线"),
        // [10] 11pt bold colour=0x38
        dx(0x00DC, 0x0038, 0x02BC, 0x0001, "等线"),
        // [11] 12pt colour=0x14
        dx(0x00F0, 0x0014, 0x0190, 0x0000, "等线"),
        // [12] 12pt colour=0x11
        dx(0x00F0, 0x0011, 0x0190, 0x0000, "等线"),
        // [13] 12pt bold colour=0x08
        dx(0x00F0, 0x0008, 0x02BC, 0x0001, "等线"),
        // [14] 12pt bold colour=0x34
        dx(0x00F0, 0x0034, 0x02BC, 0x0001, "等线"),
        // [15] 12pt bold colour=0x09
        dx(0x00F0, 0x0009, 0x02BC, 0x0001, "等线"),
        // [16] 12pt italic colour=0x17
        dx(0x00F0, 0x0017, 0x0190, 0x0002, "等线"),
        // [17] 12pt colour=0x0A
        dx(0x00F0, 0x000A, 0x0190, 0x0000, "等线"),
        // [18] 12pt colour=0x34
        dx(0x00F0, 0x0034, 0x0190, 0x0000, "等线"),
        // [19] 12pt colour=0x3C
        dx(0x00F0, 0x003C, 0x0190, 0x0000, "等线"),
        // [20] 12pt bold colour=0x3F
        dx(0x00F0, 0x003F, 0x02BC, 0x0001, "等线"),
        // [21] 12pt colour=0x3E
        dx(0x00F0, 0x003E, 0x0190, 0x0000, "等线"),
        // [22] 12pt colour=0x09
        dx(0x00F0, 0x0009, 0x0190, 0x0000, "等线"),
    ];
    for f in &fonts {
        buf.extend_from_slice(&f.serialize());
    }
    buf
}

pub fn serialize_default_formats() -> Vec<u8> {
    let mut buf = Vec::new();
    let formats = [
        (5, "\"¥\"#,##0_);\\(\"¥\"#,##0\\)"),
        (6, "\"¥\"#,##0_);[Red]\\(\"¥\"#,##0\\)"),
        (7, "\"¥\"#,##0.00_);\\(\"¥\"#,##0.00\\)"),
        (8, "\"¥\"#,##0.00_);[Red]\\(\"¥\"#,##0.00\\)"),
        (42, "_(\"¥\"* #,##0_);_(\"¥\"* \\(#,##0\\);_(\"¥\"* \"-\"_);_(@_)"),
        (41, "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)"),
        (44, "_(\"¥\"* #,##0.00_);_(\"¥\"* \\(#,##0.00\\);_(\"¥\"* \"-\"??_);_(@_)"),
        (43, "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)"),
        (23, "\\$#,##0_);\\(\\$#,##0\\)"),
        (24, "\\$#,##0_);[Red]\\(\\$#,##0\\)"),
        (25, "\\$#,##0.00_);\\(\\$#,##0.00\\)"),
        (26, "\\$#,##0.00_);[Red]\\(\\$#,##0.00\\)"),
    ];
    for (ifmt, s) in &formats {
        let rec = NumberFormatRecord::new(*ifmt, s);
        buf.extend_from_slice(&rec.serialize());
    }
    buf
}

pub fn serialize_default_xf_records() -> Vec<u8> {
    let mut buf = Vec::new();
    let xf_records: [XFRecord; 62] = [
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, used_attributes: 0x00, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 1, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 1, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 2, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 2, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 0, alignment: Alignment { vert: 1, ..Default::default() }, used_attributes: 0x00, ..Default::default() }, xf_type: XFType::Cell, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 27, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 47, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 42, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 27, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 45, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 42, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 44, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 47, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 49, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 44, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 45, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 42, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 49, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 47, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 57, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 48, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 25, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 7, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 11, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 5, format_idx: 9, alignment: Alignment { vert: 1, ..Default::default() }, used_attributes: 0xf8, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 8, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 9, alignment: Alignment { vert: 1, ..Default::default() }, borders: Borders { bottom: 5, bottom_colour: 21, ..Default::default() }, used_attributes: 0xd4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 10, alignment: Alignment { vert: 1, ..Default::default() }, borders: Borders { bottom: 5, bottom_colour: 49, ..Default::default() }, used_attributes: 0xd4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 11, alignment: Alignment { vert: 1, ..Default::default() }, borders: Borders { bottom: 2, bottom_colour: 49, ..Default::default() }, used_attributes: 0xd4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 11, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 12, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 45, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 13, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 42, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 14, alignment: Alignment { vert: 1, ..Default::default() }, borders: Borders { top: 1, bottom: 6, top_colour: 21, bottom_colour: 21, ..Default::default() }, used_attributes: 0xd4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 5, format_idx: 44, alignment: Alignment { vert: 1, ..Default::default() }, used_attributes: 0xf8, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 5, format_idx: 42, alignment: Alignment { vert: 1, ..Default::default() }, used_attributes: 0xf8, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 15, alignment: Alignment { vert: 1, ..Default::default() }, borders: Borders { left: 1, right: 1, top: 1, bottom: 1, left_colour: 23, right_colour: 23, top_colour: 23, bottom_colour: 23, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 22, ..Default::default() }, used_attributes: 0x94, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 16, alignment: Alignment { vert: 1, ..Default::default() }, borders: Borders { left: 6, right: 6, top: 6, bottom: 6, left_colour: 63, right_colour: 63, top_colour: 63, bottom_colour: 63, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 55, ..Default::default() }, used_attributes: 0x94, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 17, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 18, alignment: Alignment { vert: 1, ..Default::default() }, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 19, alignment: Alignment { vert: 1, ..Default::default() }, borders: Borders { bottom: 6, bottom_colour: 52, ..Default::default() }, used_attributes: 0xd4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 5, format_idx: 43, alignment: Alignment { vert: 1, ..Default::default() }, used_attributes: 0xf8, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 5, format_idx: 41, alignment: Alignment { vert: 1, ..Default::default() }, used_attributes: 0xf8, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 20, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 43, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 21, alignment: Alignment { vert: 1, ..Default::default() }, borders: Borders { left: 1, right: 1, top: 1, bottom: 1, left_colour: 63, right_colour: 63, top_colour: 63, bottom_colour: 63, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 22, ..Default::default() }, used_attributes: 0x94, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 22, alignment: Alignment { vert: 1, ..Default::default() }, borders: Borders { left: 1, right: 1, top: 1, bottom: 1, left_colour: 23, right_colour: 23, top_colour: 23, bottom_colour: 23, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 47, ..Default::default() }, used_attributes: 0x94, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 23, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 21, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 23, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 53, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 23, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 57, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 23, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 40, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 23, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 25, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 23, alignment: Alignment { vert: 1, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 57, ..Default::default() }, used_attributes: 0xb4, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
        XFRecord { xf: XF { font_idx: 5, alignment: Alignment { vert: 1, ..Default::default() }, borders: Borders { left: 1, right: 1, top: 1, bottom: 1, left_colour: 22, right_colour: 22, top_colour: 22, bottom_colour: 22, ..Default::default() }, pattern: Pattern { pattern: 1, pattern_fore_colour: 26, ..Default::default() }, used_attributes: 0x9c, ..Default::default() }, xf_type: XFType::Style, parent_index: 0 },
    ];
    for rec in &xf_records {
        buf.extend_from_slice(&rec.serialize());
    }
    buf
}

// ===== BIFF8 Extension Records =====
pub const EXT_087C_0: HexBiffRecord = HexBiffRecord::new(0x087C, "7c080000000000000000000000003e0017a3bf8a");

pub fn serialize_ext_087c() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_087C_0.serialize());
    buf
}

pub const EXT_087D_0: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000000000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_1: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000500000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_2: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000600000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_3: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000700000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_4: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000800000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_5: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000900000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_6: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000a00000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_7: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000b00000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_8: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000c00000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_9: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000d00000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_10: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000e00000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_11: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000000f00000002000d00140003000000010000002e30305c295f282a0e00050002");
pub const EXT_087D_12: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001000000003000d00140003000000010000002e30305c295f282a0e000500020400140003006566040000003b5f28405f292020");
pub const EXT_087D_13: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001100000003000d00140003000000010000002e30305c295f282a0e000500020400140003006566050000003b5f28405f292020");
pub const EXT_087D_14: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001200000003000d00140003000000010000002e30305c295f282a0e000500020400140003006566060000003b5f28405f292020");
pub const EXT_087D_15: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001300000003000d00140003000000010000002e30305c295f282a0e000500020400140003006566070000003b5f28405f292020");
pub const EXT_087D_16: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001400000003000d00140003000000010000002e30305c295f282a0e000500020400140003006566080000003b5f28405f292020");
pub const EXT_087D_17: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001500000003000d00140003000000010000002e30305c295f282a0e000500020400140003006566090000003b5f28405f292020");
pub const EXT_087D_18: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001600000003000d00140003000000010000002e30305c295f282a0e00050002040014000300cc4c040000003b5f28405f292020");
pub const EXT_087D_19: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001700000003000d00140003000000010000002e30305c295f282a0e00050002040014000300cc4c050000003b5f28405f292020");
pub const EXT_087D_20: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001800000003000d00140003000000010000002e30305c295f282a0e00050002040014000300cc4c060000003b5f28405f292020");
pub const EXT_087D_21: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001900000003000d00140003000000010000002e30305c295f282a0e00050002040014000300cc4c070000003b5f28405f292020");
pub const EXT_087D_22: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001a00000003000d00140003000000010000002e30305c295f282a0e00050002040014000300cc4c080000003b5f28405f292020");
pub const EXT_087D_23: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001b00000003000d00140003000000010000002e30305c295f282a0e00050002040014000300cc4c090000003b5f28405f292020");
pub const EXT_087D_24: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001c00000003000d00140003000000010000002e30305c295f282a0e000500020400140003003233040000003b5f28405f292020");
pub const EXT_087D_25: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001d00000003000d00140003000000010000002e30305c295f282a0e000500020400140003003233050000003b5f28405f292020");
pub const EXT_087D_26: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001e00000003000d00140003000000010000002e30305c295f282a0e000500020400140003003233060000003b5f28405f292020");
pub const EXT_087D_27: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000001f00000003000d00140003000000010000002e30305c295f282a0e000500020400140003003233070000003b5f28405f292020");
pub const EXT_087D_28: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002000000003000d00140003000000010000002e30305c295f282a0e000500020400140003003233080000003b5f28405f292020");
pub const EXT_087D_29: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002100000003000d00140003000000010000002e30305c295f282a0e000500020400140003003233090000003b5f28405f292020");
pub const EXT_087D_30: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002300000002000d00140003000000030000002e30305c295f282a0e00050001");
pub const EXT_087D_31: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002400000003000d00140003000000030000002e30305c295f282a0e000500020800140003000000040000003b5f28405f292020");
pub const EXT_087D_32: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002500000003000d00140003000000030000002e30305c295f282a0e00050002080014000300ff3f040000003b5f28405f292020");
pub const EXT_087D_33: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002600000003000d00140003000000030000002e30305c295f282a0e000500020800140003003233040000003b5f28405f292020");
pub const EXT_087D_34: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002700000002000d00140003000000030000002e30305c295f282a0e00050002");
pub const EXT_087D_35: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002800000003000d001400020000009c0006ff2e30305c295f282a0e000500020400140002000000ffc7ceff3b5f28405f292020");
pub const EXT_087D_36: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002900000003000d00140002000000006100ff2e30305c295f282a0e000500020400140002000000c6efceff3b5f28405f292020");
pub const EXT_087D_37: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002a00000004000d00140003000000010000002e30305c295f282a0e000500020700140003000000040000003b5f28405f2920200800140003000000040000002020202020202020");
pub const EXT_087D_38: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002d00000007000d00140002000000fa7d00ff2e30305c295f282a0e000500020400140002000000f2f2f2ff3b5f28405f29202007001400020000007f7f7fff202020202020202008001400020000007f7f7fff202020202020202009001400020000007f7f7fff00000000000000000a001400020000007f7f7fff0000000000000000");
pub const EXT_087D_39: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002e00000007000d00140003000000000000002e30305c295f282a0e000500020400140002000000a5a5a5ff3b5f28405f29202007001400020000003f3f3fff202020202020202008001400020000003f3f3fff202020202020202009001400020000003f3f3fff00000000000000000a001400020000003f3f3fff0000000000000000");
pub const EXT_087D_40: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000002f00000002000d001400020000007f7f7fff2e30305c295f282a0e00050002");
pub const EXT_087D_41: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003000000002000d00140002000000ff0000ff2e30305c295f282a0e00050002");
pub const EXT_087D_42: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003100000003000d00140002000000fa7d00ff2e30305c295f282a0e000500020800140002000000ff8001ff3b5f28405f292020");
pub const EXT_087D_43: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003400000003000d001400020000009c5700ff2e30305c295f282a0e000500020400140002000000ffeb9cff3b5f28405f292020");
pub const EXT_087D_44: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003500000007000d001400020000003f3f3fff2e30305c295f282a0e000500020400140002000000f2f2f2ff3b5f28405f29202007001400020000003f3f3fff202020202020202008001400020000003f3f3fff202020202020202009001400020000003f3f3fff00000000000000000a001400020000003f3f3fff0000000000000000");
pub const EXT_087D_45: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003600000007000d001400020000003f3f76ff2e30305c295f282a0e000500020400140002000000ffcc99ff3b5f28405f29202007001400020000007f7f7fff202020202020202008001400020000007f7f7fff202020202020202009001400020000007f7f7fff00000000000000000a001400020000007f7f7fff0000000000000000");
pub const EXT_087D_46: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003700000003000d00140003000000000000002e30305c295f282a0e000500020400140003000000040000003b5f28405f292020");
pub const EXT_087D_47: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003800000003000d00140003000000000000002e30305c295f282a0e000500020400140003000000050000003b5f28405f292020");
pub const EXT_087D_48: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003900000003000d00140003000000000000002e30305c295f282a0e000500020400140003000000060000003b5f28405f292020");
pub const EXT_087D_49: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003a00000003000d00140003000000000000002e30305c295f282a0e000500020400140003000000070000003b5f28405f292020");
pub const EXT_087D_50: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003b00000003000d00140003000000000000002e30305c295f282a0e000500020400140003000000080000003b5f28405f292020");
pub const EXT_087D_51: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003c00000003000d00140003000000000000002e30305c295f282a0e000500020400140003000000090000003b5f28405f292020");
pub const EXT_087D_52: HexBiffRecord = HexBiffRecord::new(0x087D, "7d080000000000000000000000003d00000005000400140002000000ffffccff2e30305c295f282a0700140002000000b2b2b2ff00090000003b5f280800140002000000b2b2b2ff007f7f7fff2020200900140002000000b2b2b2ff007f7f7fff2020200a00140002000000b2b2b2ff007f7f7fff000000");

pub fn serialize_ext_087d() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_087D_0.serialize());
    buf.extend_from_slice(&EXT_087D_1.serialize());
    buf.extend_from_slice(&EXT_087D_2.serialize());
    buf.extend_from_slice(&EXT_087D_3.serialize());
    buf.extend_from_slice(&EXT_087D_4.serialize());
    buf.extend_from_slice(&EXT_087D_5.serialize());
    buf.extend_from_slice(&EXT_087D_6.serialize());
    buf.extend_from_slice(&EXT_087D_7.serialize());
    buf.extend_from_slice(&EXT_087D_8.serialize());
    buf.extend_from_slice(&EXT_087D_9.serialize());
    buf.extend_from_slice(&EXT_087D_10.serialize());
    buf.extend_from_slice(&EXT_087D_11.serialize());
    buf.extend_from_slice(&EXT_087D_12.serialize());
    buf.extend_from_slice(&EXT_087D_13.serialize());
    buf.extend_from_slice(&EXT_087D_14.serialize());
    buf.extend_from_slice(&EXT_087D_15.serialize());
    buf.extend_from_slice(&EXT_087D_16.serialize());
    buf.extend_from_slice(&EXT_087D_17.serialize());
    buf.extend_from_slice(&EXT_087D_18.serialize());
    buf.extend_from_slice(&EXT_087D_19.serialize());
    buf.extend_from_slice(&EXT_087D_20.serialize());
    buf.extend_from_slice(&EXT_087D_21.serialize());
    buf.extend_from_slice(&EXT_087D_22.serialize());
    buf.extend_from_slice(&EXT_087D_23.serialize());
    buf.extend_from_slice(&EXT_087D_24.serialize());
    buf.extend_from_slice(&EXT_087D_25.serialize());
    buf.extend_from_slice(&EXT_087D_26.serialize());
    buf.extend_from_slice(&EXT_087D_27.serialize());
    buf.extend_from_slice(&EXT_087D_28.serialize());
    buf.extend_from_slice(&EXT_087D_29.serialize());
    buf.extend_from_slice(&EXT_087D_30.serialize());
    buf.extend_from_slice(&EXT_087D_31.serialize());
    buf.extend_from_slice(&EXT_087D_32.serialize());
    buf.extend_from_slice(&EXT_087D_33.serialize());
    buf.extend_from_slice(&EXT_087D_34.serialize());
    buf.extend_from_slice(&EXT_087D_35.serialize());
    buf.extend_from_slice(&EXT_087D_36.serialize());
    buf.extend_from_slice(&EXT_087D_37.serialize());
    buf.extend_from_slice(&EXT_087D_38.serialize());
    buf.extend_from_slice(&EXT_087D_39.serialize());
    buf.extend_from_slice(&EXT_087D_40.serialize());
    buf.extend_from_slice(&EXT_087D_41.serialize());
    buf.extend_from_slice(&EXT_087D_42.serialize());
    buf.extend_from_slice(&EXT_087D_43.serialize());
    buf.extend_from_slice(&EXT_087D_44.serialize());
    buf.extend_from_slice(&EXT_087D_45.serialize());
    buf.extend_from_slice(&EXT_087D_46.serialize());
    buf.extend_from_slice(&EXT_087D_47.serialize());
    buf.extend_from_slice(&EXT_087D_48.serialize());
    buf.extend_from_slice(&EXT_087D_49.serialize());
    buf.extend_from_slice(&EXT_087D_50.serialize());
    buf.extend_from_slice(&EXT_087D_51.serialize());
    buf.extend_from_slice(&EXT_087D_52.serialize());
    buf
}

pub const EXT_0293_0: HexBiffRecord = HexBiffRecord::new(0x0293, "10000a000132003000250020002d0020004077728220003100");
pub const EXT_0293_1: HexBiffRecord = HexBiffRecord::new(0x0293, "11000a000132003000250020002d0020004077728220003200");
pub const EXT_0293_2: HexBiffRecord = HexBiffRecord::new(0x0293, "12000a000132003000250020002d0020004077728220003300");
pub const EXT_0293_3: HexBiffRecord = HexBiffRecord::new(0x0293, "13000a000132003000250020002d0020004077728220003400");
pub const EXT_0293_4: HexBiffRecord = HexBiffRecord::new(0x0293, "14000a000132003000250020002d0020004077728220003500");
pub const EXT_0293_5: HexBiffRecord = HexBiffRecord::new(0x0293, "15000a000132003000250020002d0020004077728220003600");
pub const EXT_0293_6: HexBiffRecord = HexBiffRecord::new(0x0293, "16000a000134003000250020002d0020004077728220003100");
pub const EXT_0293_7: HexBiffRecord = HexBiffRecord::new(0x0293, "17000a000134003000250020002d0020004077728220003200");
pub const EXT_0293_8: HexBiffRecord = HexBiffRecord::new(0x0293, "18000a000134003000250020002d0020004077728220003300");
pub const EXT_0293_9: HexBiffRecord = HexBiffRecord::new(0x0293, "19000a000134003000250020002d0020004077728220003400");
pub const EXT_0293_10: HexBiffRecord = HexBiffRecord::new(0x0293, "1a000a000134003000250020002d0020004077728220003500");
pub const EXT_0293_11: HexBiffRecord = HexBiffRecord::new(0x0293, "1b000a000134003000250020002d0020004077728220003600");
pub const EXT_0293_12: HexBiffRecord = HexBiffRecord::new(0x0293, "1c000a000136003000250020002d0020004077728220003100");
pub const EXT_0293_13: HexBiffRecord = HexBiffRecord::new(0x0293, "1d000a000136003000250020002d0020004077728220003200");
pub const EXT_0293_14: HexBiffRecord = HexBiffRecord::new(0x0293, "1e000a000136003000250020002d0020004077728220003300");
pub const EXT_0293_15: HexBiffRecord = HexBiffRecord::new(0x0293, "1f000a000136003000250020002d0020004077728220003400");
pub const EXT_0293_16: HexBiffRecord = HexBiffRecord::new(0x0293, "20000a000136003000250020002d0020004077728220003500");
pub const EXT_0293_17: HexBiffRecord = HexBiffRecord::new(0x0293, "21000a000136003000250020002d0020004077728220003600");
pub const EXT_0293_18: HexBiffRecord = HexBiffRecord::new(0x0293, "228005ff");
pub const EXT_0293_19: HexBiffRecord = HexBiffRecord::new(0x0293, "230002000107689898");
pub const EXT_0293_20: HexBiffRecord = HexBiffRecord::new(0x0293, "24000400010768989820003100");
pub const EXT_0293_21: HexBiffRecord = HexBiffRecord::new(0x0293, "25000400010768989820003200");
pub const EXT_0293_22: HexBiffRecord = HexBiffRecord::new(0x0293, "26000400010768989820003300");
pub const EXT_0293_23: HexBiffRecord = HexBiffRecord::new(0x0293, "27000400010768989820003400");
pub const EXT_0293_24: HexBiffRecord = HexBiffRecord::new(0x0293, "2800010001ee5d");
pub const EXT_0293_25: HexBiffRecord = HexBiffRecord::new(0x0293, "008000ff");
pub const EXT_0293_26: HexBiffRecord = HexBiffRecord::new(0x0293, "29000100017d59");
pub const EXT_0293_27: HexBiffRecord = HexBiffRecord::new(0x0293, "2a00020001476c3b60");
pub const EXT_0293_28: HexBiffRecord = HexBiffRecord::new(0x0293, "2b8004ff");
pub const EXT_0293_29: HexBiffRecord = HexBiffRecord::new(0x0293, "2c8007ff");
pub const EXT_0293_30: HexBiffRecord = HexBiffRecord::new(0x0293, "2d00020001a18b977b");
pub const EXT_0293_31: HexBiffRecord = HexBiffRecord::new(0x0293, "2e00050001c068e567555343513c68");
pub const EXT_0293_32: HexBiffRecord = HexBiffRecord::new(0x0293, "2f00050001e389ca91276087652c67");
pub const EXT_0293_33: HexBiffRecord = HexBiffRecord::new(0x0293, "3000040001668b4a5487652c67");
pub const EXT_0293_34: HexBiffRecord = HexBiffRecord::new(0x0293, "3100050001fe94a563555343513c68");
pub const EXT_0293_35: HexBiffRecord = HexBiffRecord::new(0x0293, "328003ff");
pub const EXT_0293_36: HexBiffRecord = HexBiffRecord::new(0x0293, "338006ff");
pub const EXT_0293_37: HexBiffRecord = HexBiffRecord::new(0x0293, "340002000102902d4e");
pub const EXT_0293_38: HexBiffRecord = HexBiffRecord::new(0x0293, "3500020001938ffa51");
pub const EXT_0293_39: HexBiffRecord = HexBiffRecord::new(0x0293, "3600020001938f6551");
pub const EXT_0293_40: HexBiffRecord = HexBiffRecord::new(0x0293, "37000400014077728220003100");
pub const EXT_0293_41: HexBiffRecord = HexBiffRecord::new(0x0293, "38000400014077728220003200");
pub const EXT_0293_42: HexBiffRecord = HexBiffRecord::new(0x0293, "39000400014077728220003300");
pub const EXT_0293_43: HexBiffRecord = HexBiffRecord::new(0x0293, "3a000400014077728220003400");
pub const EXT_0293_44: HexBiffRecord = HexBiffRecord::new(0x0293, "3b000400014077728220003500");
pub const EXT_0293_45: HexBiffRecord = HexBiffRecord::new(0x0293, "3c000400014077728220003600");
pub const EXT_0293_46: HexBiffRecord = HexBiffRecord::new(0x0293, "3d00020001e86cca91");

pub fn serialize_ext_0293() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_0293_0.serialize());
    buf.extend_from_slice(&EXT_0293_1.serialize());
    buf.extend_from_slice(&EXT_0293_2.serialize());
    buf.extend_from_slice(&EXT_0293_3.serialize());
    buf.extend_from_slice(&EXT_0293_4.serialize());
    buf.extend_from_slice(&EXT_0293_5.serialize());
    buf.extend_from_slice(&EXT_0293_6.serialize());
    buf.extend_from_slice(&EXT_0293_7.serialize());
    buf.extend_from_slice(&EXT_0293_8.serialize());
    buf.extend_from_slice(&EXT_0293_9.serialize());
    buf.extend_from_slice(&EXT_0293_10.serialize());
    buf.extend_from_slice(&EXT_0293_11.serialize());
    buf.extend_from_slice(&EXT_0293_12.serialize());
    buf.extend_from_slice(&EXT_0293_13.serialize());
    buf.extend_from_slice(&EXT_0293_14.serialize());
    buf.extend_from_slice(&EXT_0293_15.serialize());
    buf.extend_from_slice(&EXT_0293_16.serialize());
    buf.extend_from_slice(&EXT_0293_17.serialize());
    buf.extend_from_slice(&EXT_0293_18.serialize());
    buf.extend_from_slice(&EXT_0293_19.serialize());
    buf.extend_from_slice(&EXT_0293_20.serialize());
    buf.extend_from_slice(&EXT_0293_21.serialize());
    buf.extend_from_slice(&EXT_0293_22.serialize());
    buf.extend_from_slice(&EXT_0293_23.serialize());
    buf.extend_from_slice(&EXT_0293_24.serialize());
    buf.extend_from_slice(&EXT_0293_25.serialize());
    buf.extend_from_slice(&EXT_0293_26.serialize());
    buf.extend_from_slice(&EXT_0293_27.serialize());
    buf.extend_from_slice(&EXT_0293_28.serialize());
    buf.extend_from_slice(&EXT_0293_29.serialize());
    buf.extend_from_slice(&EXT_0293_30.serialize());
    buf.extend_from_slice(&EXT_0293_31.serialize());
    buf.extend_from_slice(&EXT_0293_32.serialize());
    buf.extend_from_slice(&EXT_0293_33.serialize());
    buf.extend_from_slice(&EXT_0293_34.serialize());
    buf.extend_from_slice(&EXT_0293_35.serialize());
    buf.extend_from_slice(&EXT_0293_36.serialize());
    buf.extend_from_slice(&EXT_0293_37.serialize());
    buf.extend_from_slice(&EXT_0293_38.serialize());
    buf.extend_from_slice(&EXT_0293_39.serialize());
    buf.extend_from_slice(&EXT_0293_40.serialize());
    buf.extend_from_slice(&EXT_0293_41.serialize());
    buf.extend_from_slice(&EXT_0293_42.serialize());
    buf.extend_from_slice(&EXT_0293_43.serialize());
    buf.extend_from_slice(&EXT_0293_44.serialize());
    buf.extend_from_slice(&EXT_0293_45.serialize());
    buf.extend_from_slice(&EXT_0293_46.serialize());
    buf
}

pub const EXT_0892_0: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001041eff0a0032003000250020002d00200040777282200031000000030001000c0007046566c0e6f5ff05000c0007010000000000ff2500050002");
pub const EXT_0892_1: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010422ff0a0032003000250020002d00200040777282200032000000030001000c0007056566fbe2d5ff05000c0007010000000000ff2500050002");
pub const EXT_0892_2: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010426ff0a0032003000250020002d00200040777282200033000000030001000c0007066566c1f0c8ff05000c0007010000000000ff2500050002");
pub const EXT_0892_3: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001042aff0a0032003000250020002d00200040777282200034000000030001000c0007076566caedfbff05000c0007010000000000ff2500050002");
pub const EXT_0892_4: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001042eff0a0032003000250020002d00200040777282200035000000030001000c0007086566f2ceefff05000c0007010000000000ff2500050002");
pub const EXT_0892_5: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010432ff0a0032003000250020002d00200040777282200036000000030001000c0007096566daf2d0ff05000c0007010000000000ff2500050002");
pub const EXT_0892_6: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001041fff0a0034003000250020002d00200040777282200031000000030001000c000704cc4c83ccebff05000c0007010000000000ff2500050002");
pub const EXT_0892_7: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010423ff0a0034003000250020002d00200040777282200032000000030001000c000705cc4cf7c7acff05000c0007010000000000ff2500050002");
pub const EXT_0892_8: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010427ff0a0034003000250020002d00200040777282200033000000030001000c000706cc4c83e28eff05000c0007010000000000ff2500050002");
pub const EXT_0892_9: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001042bff0a0034003000250020002d00200040777282200034000000030001000c000707cc4c94dcf8ff05000c0007010000000000ff2500050002");
pub const EXT_0892_10: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001042fff0a0034003000250020002d00200040777282200035000000030001000c000708cc4ce49eddff05000c0007010000000000ff2500050002");
pub const EXT_0892_11: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010433ff0a0034003000250020002d00200040777282200036000000030001000c000709cc4cb5e6a2ff05000c0007010000000000ff2500050002");
pub const EXT_0892_12: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010420ff0a0036003000250020002d00200040777282200031000000030001000c000704323344b3e1ff05000c0007010000000000ff2500050002");
pub const EXT_0892_13: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010424ff0a0036003000250020002d00200040777282200032000000030001000c0007053233f1a983ff05000c0007010000000000ff2500050002");
pub const EXT_0892_14: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010428ff0a0036003000250020002d00200040777282200033000000030001000c000706323347d359ff05000c0007010000000000ff2500050002");
pub const EXT_0892_15: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001042cff0a0036003000250020002d00200040777282200034000000030001000c000707323361cbf3ff05000c0007010000000000ff2500050002");
pub const EXT_0892_16: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010430ff0a0036003000250020002d00200040777282200035000000030001000c0007083233d86dcdff05000c0007010000000000ff2500050002");
pub const EXT_0892_17: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010434ff0a0036003000250020002d00200040777282200036000000030001000c00070932338ed973ff05000c0007010000000000ff2500050002");
pub const EXT_0892_18: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010505ff03007e760652d46b00000000");
pub const EXT_0892_19: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001030fff0200076898980000020005000c00070300000e2841ff2500050001");
pub const EXT_0892_20: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010310ff040007689898200031000000030005000c00070300000e2841ff250005000207000e0007040000156082ff0500");
pub const EXT_0892_21: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010311ff040007689898200032000000030005000c00070300000e2841ff250005000207000e000704ff3f64bee6ff0500");
pub const EXT_0892_22: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010312ff040007689898200033000000030005000c00070300000e2841ff250005000207000e000704323344b3e1ff0200");
pub const EXT_0892_23: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010313ff040007689898200034000000020005000c00070300000e2841ff2500050002");
pub const EXT_0892_24: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001011bff0100ee5d0000030001000c0005ff0000ffc7ceff05000c0005ff00009c0006ff2500050002");
pub const EXT_0892_25: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010100ff0200385ec4890000020005000c0007010000000000ff2500050002");
pub const EXT_0892_26: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001011aff01007d590000030001000c0005ff0000c6efceff05000c0005ff0000006100ff2500050002");
pub const EXT_0892_27: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010319ff0200476c3b600000040005000c0007010000000000ff250005000206000e0007040000156082ff010007000e0007040000156082ff0600");
pub const EXT_0892_28: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010504ff0200278d015e00000000");
pub const EXT_0892_29: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010507ff0500278d015e5b0030005d0000000000");
pub const EXT_0892_30: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010216ff0200a18b977b0000070001000c0005ff0000f2f2f2ff05000c0005ff0000fa7d00ff250005000206000e0005ff00007f7f7fff010007000e0005ff00007f7f7fff010008000e0005ff00007f7f7fff010009000e0005ff00007f7f7fff0100");
pub const EXT_0892_31: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010217ff0500c068e567555343513c680000070001000c0005ff0000a5a5a5ff05000c0007000000ffffffff250005000206000e0005ff00003f3f3fff060007000e0005ff00003f3f3fff060008000e0005ff00003f3f3fff060009000e0005ff00003f3f3fff0600");
pub const EXT_0892_32: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010235ff0500e389ca91276087652c670000020005000c0005ff00007f7f7fff2500050002");
pub const EXT_0892_33: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001020bff0400668b4a5487652c670000020005000c0005ff0000ff0000ff2500050002");
pub const EXT_0892_34: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010218ff0500fe94a563555343513c680000030005000c0005ff0000fa7d00ff250005000207000e0005ff0000ff8001ff0600");
pub const EXT_0892_35: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010503ff040043534d4f0652949600000000");
pub const EXT_0892_36: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010506ff070043534d4f065294965b0030005d0000000000");
pub const EXT_0892_37: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001011cff020002902d4e0000030001000c0005ff0000ffeb9cff05000c0005ff00009c5700ff2500050002");
pub const EXT_0892_38: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010215ff0200938ffa510000070001000c0005ff0000f2f2f2ff05000c0005ff00003f3f3fff250005000206000e0005ff00003f3f3fff010007000e0005ff00003f3f3fff010008000e0005ff00003f3f3fff010009000e0005ff00003f3f3fff0100");
pub const EXT_0892_39: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010214ff0200938f65510000070001000c0005ff0000ffcc99ff05000c0005ff00003f3f76ff250005000206000e0005ff00007f7f7fff010007000e0005ff00007f7f7fff010008000e0005ff00007f7f7fff010009000e0005ff00007f7f7fff0100");
pub const EXT_0892_40: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001041dff040040777282200031000000030001000c0007040000156082ff05000c0007000000ffffffff2500050002");
pub const EXT_0892_41: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010421ff040040777282200032000000030001000c0007050000e97132ff05000c0007000000ffffffff2500050002");
pub const EXT_0892_42: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010425ff040040777282200033000000030001000c0007060000196b24ff05000c0007000000ffffffff2500050002");
pub const EXT_0892_43: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010429ff040040777282200034000000030001000c00070700000f9ed5ff05000c0007000000ffffffff2500050002");
pub const EXT_0892_44: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001042dff040040777282200035000000030001000c0007080000a02b93ff05000c0007000000ffffffff2500050002");
pub const EXT_0892_45: HexBiffRecord = HexBiffRecord::new(0x0892, "920800000000000000000000010431ff040040777282200036000000030001000c00070900004ea72eff05000c0007000000ffffffff2500050002");
pub const EXT_0892_46: HexBiffRecord = HexBiffRecord::new(0x0892, "92080000000000000000000001020aff0200e86cca910000050001000c0005ff0000ffffccff06000e0005ff0000b2b2b2ff010007000e0005ff0000b2b2b2ff010008000e0005ff0000b2b2b2ff010009000e0005ff0000b2b2b2ff0100");

pub fn serialize_ext_0892() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_0892_0.serialize());
    buf.extend_from_slice(&EXT_0892_1.serialize());
    buf.extend_from_slice(&EXT_0892_2.serialize());
    buf.extend_from_slice(&EXT_0892_3.serialize());
    buf.extend_from_slice(&EXT_0892_4.serialize());
    buf.extend_from_slice(&EXT_0892_5.serialize());
    buf.extend_from_slice(&EXT_0892_6.serialize());
    buf.extend_from_slice(&EXT_0892_7.serialize());
    buf.extend_from_slice(&EXT_0892_8.serialize());
    buf.extend_from_slice(&EXT_0892_9.serialize());
    buf.extend_from_slice(&EXT_0892_10.serialize());
    buf.extend_from_slice(&EXT_0892_11.serialize());
    buf.extend_from_slice(&EXT_0892_12.serialize());
    buf.extend_from_slice(&EXT_0892_13.serialize());
    buf.extend_from_slice(&EXT_0892_14.serialize());
    buf.extend_from_slice(&EXT_0892_15.serialize());
    buf.extend_from_slice(&EXT_0892_16.serialize());
    buf.extend_from_slice(&EXT_0892_17.serialize());
    buf.extend_from_slice(&EXT_0892_18.serialize());
    buf.extend_from_slice(&EXT_0892_19.serialize());
    buf.extend_from_slice(&EXT_0892_20.serialize());
    buf.extend_from_slice(&EXT_0892_21.serialize());
    buf.extend_from_slice(&EXT_0892_22.serialize());
    buf.extend_from_slice(&EXT_0892_23.serialize());
    buf.extend_from_slice(&EXT_0892_24.serialize());
    buf.extend_from_slice(&EXT_0892_25.serialize());
    buf.extend_from_slice(&EXT_0892_26.serialize());
    buf.extend_from_slice(&EXT_0892_27.serialize());
    buf.extend_from_slice(&EXT_0892_28.serialize());
    buf.extend_from_slice(&EXT_0892_29.serialize());
    buf.extend_from_slice(&EXT_0892_30.serialize());
    buf.extend_from_slice(&EXT_0892_31.serialize());
    buf.extend_from_slice(&EXT_0892_32.serialize());
    buf.extend_from_slice(&EXT_0892_33.serialize());
    buf.extend_from_slice(&EXT_0892_34.serialize());
    buf.extend_from_slice(&EXT_0892_35.serialize());
    buf.extend_from_slice(&EXT_0892_36.serialize());
    buf.extend_from_slice(&EXT_0892_37.serialize());
    buf.extend_from_slice(&EXT_0892_38.serialize());
    buf.extend_from_slice(&EXT_0892_39.serialize());
    buf.extend_from_slice(&EXT_0892_40.serialize());
    buf.extend_from_slice(&EXT_0892_41.serialize());
    buf.extend_from_slice(&EXT_0892_42.serialize());
    buf.extend_from_slice(&EXT_0892_43.serialize());
    buf.extend_from_slice(&EXT_0892_44.serialize());
    buf.extend_from_slice(&EXT_0892_45.serialize());
    buf.extend_from_slice(&EXT_0892_46.serialize());
    buf
}

/// 0293 与 0892 交替配对串行化（47 对）
pub fn serialize_ext_0293_0892_interleaved() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_0293_0.serialize());
    buf.extend_from_slice(&EXT_0892_0.serialize());
    buf.extend_from_slice(&EXT_0293_1.serialize());
    buf.extend_from_slice(&EXT_0892_1.serialize());
    buf.extend_from_slice(&EXT_0293_2.serialize());
    buf.extend_from_slice(&EXT_0892_2.serialize());
    buf.extend_from_slice(&EXT_0293_3.serialize());
    buf.extend_from_slice(&EXT_0892_3.serialize());
    buf.extend_from_slice(&EXT_0293_4.serialize());
    buf.extend_from_slice(&EXT_0892_4.serialize());
    buf.extend_from_slice(&EXT_0293_5.serialize());
    buf.extend_from_slice(&EXT_0892_5.serialize());
    buf.extend_from_slice(&EXT_0293_6.serialize());
    buf.extend_from_slice(&EXT_0892_6.serialize());
    buf.extend_from_slice(&EXT_0293_7.serialize());
    buf.extend_from_slice(&EXT_0892_7.serialize());
    buf.extend_from_slice(&EXT_0293_8.serialize());
    buf.extend_from_slice(&EXT_0892_8.serialize());
    buf.extend_from_slice(&EXT_0293_9.serialize());
    buf.extend_from_slice(&EXT_0892_9.serialize());
    buf.extend_from_slice(&EXT_0293_10.serialize());
    buf.extend_from_slice(&EXT_0892_10.serialize());
    buf.extend_from_slice(&EXT_0293_11.serialize());
    buf.extend_from_slice(&EXT_0892_11.serialize());
    buf.extend_from_slice(&EXT_0293_12.serialize());
    buf.extend_from_slice(&EXT_0892_12.serialize());
    buf.extend_from_slice(&EXT_0293_13.serialize());
    buf.extend_from_slice(&EXT_0892_13.serialize());
    buf.extend_from_slice(&EXT_0293_14.serialize());
    buf.extend_from_slice(&EXT_0892_14.serialize());
    buf.extend_from_slice(&EXT_0293_15.serialize());
    buf.extend_from_slice(&EXT_0892_15.serialize());
    buf.extend_from_slice(&EXT_0293_16.serialize());
    buf.extend_from_slice(&EXT_0892_16.serialize());
    buf.extend_from_slice(&EXT_0293_17.serialize());
    buf.extend_from_slice(&EXT_0892_17.serialize());
    buf.extend_from_slice(&EXT_0293_18.serialize());
    buf.extend_from_slice(&EXT_0892_18.serialize());
    buf.extend_from_slice(&EXT_0293_19.serialize());
    buf.extend_from_slice(&EXT_0892_19.serialize());
    buf.extend_from_slice(&EXT_0293_20.serialize());
    buf.extend_from_slice(&EXT_0892_20.serialize());
    buf.extend_from_slice(&EXT_0293_21.serialize());
    buf.extend_from_slice(&EXT_0892_21.serialize());
    buf.extend_from_slice(&EXT_0293_22.serialize());
    buf.extend_from_slice(&EXT_0892_22.serialize());
    buf.extend_from_slice(&EXT_0293_23.serialize());
    buf.extend_from_slice(&EXT_0892_23.serialize());
    buf.extend_from_slice(&EXT_0293_24.serialize());
    buf.extend_from_slice(&EXT_0892_24.serialize());
    buf.extend_from_slice(&EXT_0293_25.serialize());
    buf.extend_from_slice(&EXT_0892_25.serialize());
    buf.extend_from_slice(&EXT_0293_26.serialize());
    buf.extend_from_slice(&EXT_0892_26.serialize());
    buf.extend_from_slice(&EXT_0293_27.serialize());
    buf.extend_from_slice(&EXT_0892_27.serialize());
    buf.extend_from_slice(&EXT_0293_28.serialize());
    buf.extend_from_slice(&EXT_0892_28.serialize());
    buf.extend_from_slice(&EXT_0293_29.serialize());
    buf.extend_from_slice(&EXT_0892_29.serialize());
    buf.extend_from_slice(&EXT_0293_30.serialize());
    buf.extend_from_slice(&EXT_0892_30.serialize());
    buf.extend_from_slice(&EXT_0293_31.serialize());
    buf.extend_from_slice(&EXT_0892_31.serialize());
    buf.extend_from_slice(&EXT_0293_32.serialize());
    buf.extend_from_slice(&EXT_0892_32.serialize());
    buf.extend_from_slice(&EXT_0293_33.serialize());
    buf.extend_from_slice(&EXT_0892_33.serialize());
    buf.extend_from_slice(&EXT_0293_34.serialize());
    buf.extend_from_slice(&EXT_0892_34.serialize());
    buf.extend_from_slice(&EXT_0293_35.serialize());
    buf.extend_from_slice(&EXT_0892_35.serialize());
    buf.extend_from_slice(&EXT_0293_36.serialize());
    buf.extend_from_slice(&EXT_0892_36.serialize());
    buf.extend_from_slice(&EXT_0293_37.serialize());
    buf.extend_from_slice(&EXT_0892_37.serialize());
    buf.extend_from_slice(&EXT_0293_38.serialize());
    buf.extend_from_slice(&EXT_0892_38.serialize());
    buf.extend_from_slice(&EXT_0293_39.serialize());
    buf.extend_from_slice(&EXT_0892_39.serialize());
    buf.extend_from_slice(&EXT_0293_40.serialize());
    buf.extend_from_slice(&EXT_0892_40.serialize());
    buf.extend_from_slice(&EXT_0293_41.serialize());
    buf.extend_from_slice(&EXT_0892_41.serialize());
    buf.extend_from_slice(&EXT_0293_42.serialize());
    buf.extend_from_slice(&EXT_0892_42.serialize());
    buf.extend_from_slice(&EXT_0293_43.serialize());
    buf.extend_from_slice(&EXT_0892_43.serialize());
    buf.extend_from_slice(&EXT_0293_44.serialize());
    buf.extend_from_slice(&EXT_0892_44.serialize());
    buf.extend_from_slice(&EXT_0293_45.serialize());
    buf.extend_from_slice(&EXT_0892_45.serialize());
    buf.extend_from_slice(&EXT_0293_46.serialize());
    buf.extend_from_slice(&EXT_0892_46.serialize());
    buf
}

pub const EXT_088E_0: HexBiffRecord = HexBiffRecord::new(0x088E, "8e080000000000000000000090000000110011005400610062006c0065005300740079006c0065004d0065006400690075006d0032005000690076006f0074005300740079006c0065004c00690067006800740031003600");

pub fn serialize_ext_088e() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_088E_0.serialize());
    buf
}

pub const EXT_0160_0: HexBiffRecord = HexBiffRecord::new(0x0160, "0000");

pub fn serialize_ext_0160() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_0160_0.serialize());
    buf
}

pub const EXT_089A_0: HexBiffRecord = HexBiffRecord::new(0x089A, "9a080000000000000000000001000000000000000a000000");

pub fn serialize_ext_089a() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_089A_0.serialize());
    buf
}

pub const EXT_08A3_0: HexBiffRecord = HexBiffRecord::new(0x08A3, "a3080000000000000000000000000000");

pub fn serialize_ext_08a3() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_08A3_0.serialize());
    buf
}

pub const EXT_008C_0: HexBiffRecord = HexBiffRecord::new(0x008C, "56005600");

pub fn serialize_ext_008c() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_008C_0.serialize());
    buf
}

pub const EXT_01C1_0: HexBiffRecord = HexBiffRecord::new(0x01C1, "c101000035ea0e00");

pub fn serialize_ext_01c1() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_01C1_0.serialize());
    buf
}

pub const EXT_0863_0: HexBiffRecord = HexBiffRecord::new(0x0863, "63080000000000000000000016000000000000000200");

pub fn serialize_ext_0863() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_0863_0.serialize());
    buf
}

pub const EXT_0896_0: HexBiffRecord = HexBiffRecord::new(0x0896, "9608000000000000000000003c160300504b030414000600080000002100e9de0fbfff0000001c020000130000005b436f6e74656e745f54797065735d2e786d6cac91cb4ec3301045f748fc83e52d4a9cb2400825e982c78ec7a27cc0c8992416c9d8b2a755fbf74cd25442a820166c2cd933f79e3be372bd1f07b5c3989ca74aaff2422b24eb1b475da5df374fd9ad5689811a183c61a50f98f4babebc2837878049899a52a57be670674cb23d8e90721f90a4d2fa3802cb35762680fd800ecd7551dc18eb899138e3c943d7e503b6b01d583deee5f99824e290b4ba3f364eac4a430883b3c092d4eca8f946c916422ecab927f52ea42b89a1cd59c254f919b0e85e6535d135a8de20f20b8c12c3b00c895fcf6720192de6bf3b9e89ecdbd6596cbcdd8eb28e7c365ecc4ec1ff1460f53fe813d3cc7f5b7f020000ffff0300504b030414000600080000002100a5d6a7e7c0000000360100000b0000005f72656c732f2e72656c73848fcf6ac3300c87ef85bd83d17d51d2c31825762fa590432fa37d00e1287f68221bdb1bebdb4fc7060abb0884a4eff7a93dfeae8bf9e194e720169aaa06c3e2433fcb68e1763dbf7f82c985a4a725085b787086a37bdbb55fbc50d1a33ccd311ba548b63095120f88d94fbc52ae4264d1c910d24a45db3462247fa791715fd71f989e19e0364cd3f51652d73760ae8fa8c9ffb3c330cc9e4fc17faf2ce545046e37944c69e462a1a82fe353bd90a865aad41ed0b5b8f9d6fd010000ffff0300504b0304140006000800000021006b799616830000008a0000001c0000007468656d652f7468656d652f7468656d654d616e616765722e786d6c0ccc4d0ac3201040e17da17790d93763bb284562b2cbaebbf600439c1a41c7a0d29fdbd7e5e38337cedf14d59b4b0d592c9c070d8a65cd2e88b7f07c2ca71ba8da481cc52c6ce1c715e6e97818c9b48d13df49c873517d23d59085adb5dd20d6b52bd521ef2cdd5eb9246a3d8b4757e8d3f729e245eb2b260a0238fd010000ffff0300504b030414000600080000002100a91cc9bc4807000002210000160000007468656d652f7468656d652f7468656d65312e786d6cec59cd6e1b3710be17e83b107b6f2cc9b66c1991033b96e324fe83a5a4c89192282d63ee724152b6750b925351a04081b4c8a540d14b0f45d1000dd0a03df45dea22419a3e4487e46a454a54fc83004d0bc73968a96f861f678633a3d9eb374e12868e889094a7f5a87cad142192767897a6fd7a74afb5f9c97284a4c26917339e927a342432bab1faf147d7f18a8a494210c8a77205d7a358a96c656e4e766019cb6b3c23297cd7e322c10a1e457fae2bf031e84dd85ca554aace2598a6114a71026af77a3dda21e8cf97bfbdf9eed91f8f3e87ffd1ea688f06838d5225f5428789a6de81788206db3d2c6b841cca9b4ca023ccea116cd7e5c72d72a222c4b054f0453d2a997fd1dceaf539bc920b313543d691db34ff72b95ca07b58317b8a7ebbd8b4d4a82c2f940bfd06c0d434aeb1acff0a7d06803b1d38a9e5e2ea2c2f564bcb951ceb80ecc780eeda5279dec73bfae7a738976bd5f5ca82a7df80acfe85297c69b3d6d858f4f00664f18b53f8b55265bd36efe10dc8e2ab53f885c6da52a5e1e10d2866343d9c4657979697ab39ba80f438db0ac26bd56a696923878f51100d4574e92d7a3c55b3622dc10fb9d804800632ac688ad430233ddc81605ecb14976883ca8ce16184329c7209cba54ab90ca1b750aa147fc6e2788560475af30226726a49f341b22368a6ead11dd01a3990572f5f9e3e7e71faf897d3274f4e1fff84b6693f56569527b785d3be2bf7f6fb2ffffee611faebe76fdf3efd2a8c972efef58f9fbdfef5f777a987ab3636c5abaf9fbf7ef1fcd5b32fdefcf034a07d4de0b60b6fd18448b44b8ed1014fe080c6143e7fd216179368c5987a123806dd01d50d157bc0dd216621dc3af14d785f409609016f0d1e7a5c9bb118281ad8f96e9c78c01dced93a174103dcd57b39166e0dd27e787331707107181f85f6be8953cfc18d4106e9958654de8c8947739fe154e13e498942fa3b7e4848e0740f28f5ecba433b824bde53e80145eb98064dd2a26d2f90c6425b3401bf0c4304c1d59e6d76eea375ce42a7de20473e12ae056601f22dc23c33dec203859390ca164e986bf06dace210c9e650745c5c432af0749f308e1a5d226548664fc0791da7dfc590d8826edf61c3c4470a450f433ab731e72e72831fde8c719285b04d9ac62ef6b63c8410c5689fab107c87fb37443f831f703ad3ddf729f1dc7d7622b80709cea5340e10fdcd40047c798b702f7e9b43d6c3249465d644e265d7354183d1b13ee87ba1bd4d08c3c7b84b08ba773bc0609d679ecdc7a4efc49055b64828b0ee603f56f5734a2441a6af994e91db547a21db247d3e83cfce7022f10c719a60314bf32e78ddb579a32de03206ceb9c73a872e7097421708f11234ca9e041d4e70cfd4ba1f63af76e967198ed7a1f0fc779e3b06f7f2a147e31cf71264c8856520b1bb32efb44d0b336f8371c0b43045dba1740b229efbc722baae1ab14150aee75fdab11ba031f2fa9d84a667353fbb58087efceff43e81787c3f5d4f58b197b72ed8efccca2b5b135dce2cdc7fb0b7d9c083749f4039994e5c57adcd556b13fdef5b9b5977f9aaa199d5765c353411341a570d4d3e5e793f0dcdb88781f6468f18eca8c70c7e9299739f1e65aca9868c6c4b33fa91f0b3a6bb098b5ace8c3e493107cc62f8a8cb1c6ce0e1fa021b1924b8fa94aab819e30ce643e5482be9cb5c755fa28c4b181b99e5a06e8d6783648777edb8d3cc974ab6b24aacc6eba545183cd9751855298bae2ee58b9a9f99a9025fc3b66f46ad23025af622249ccd7c12f301124ba3c53348e8c9d9fb61510bb058d6ea47ae9a3205502bbc02bfbb11fc5aaf478b0b9a108ccb65077af4aef69375f5c8bbc699efd3d3b38ce945008c16ed49603e5f78baa6b9ce3c9e3e9d0db57378da23619c62c3ca27612c631a3c19c3afe13c3af5ea79685cd4d7b5b14b3d7ada14663f88ef318da5e577b1b8acaf416e3237b0d4cd142c45c770c72b70e922d4c1593deac1dc183e2619048fd4bfbd30ebc33b988e12f6c65f26b56442aa0d2c636b719375ac7f12aa88408c26f5489fbff0034b4d12b1e46a70753f5472157de13e3472e075dfcba4d7231de5fadd59d196b68f90e26db2087e6bc42f0fd6927c00ee6ec6dd63d46603718021c41697cadabb5d2ae1f541d9baba4be17d5891c9c6f1375199f2ecefbe90323164d731cb629c9714379b5bb82928051df354d8c079cacf0c06754c9257c2765f5758d7a85e392d6a97e530b3ec9e2da42de764cd71d1f4d28a2e9be134e6ed30aa0313b6bc5c9577588d4c0c49cd2df136774fe6dcda28d94d340a4599008317f6bb5ced77a88d37f3a869c6d3795827ed7cd52f1ea3039e41ed3c55c249fbd591da09bb154522b81d2c5eaaf483dc64d4c2526fd4581a4b9bf7e7eebb6dde7e08c96303dadc01b36fbb590a4f3a2a65b62f8c6fdbbc3bcc3f3269138df5b96e4a3592a507a48768f7a41e55429da37de35acebb0183d662ba781582c16ecf17ccf15ad45ed842d806781154f652fac28584d9197aef42d88c1643b4d5c988b2eed5016f4c28eda9c1b4b9a5e0d3b415e1f5bfc0d0db364d6757e45e63ffd57f000000ffff0300504b0304140006000800000021000dd1909fb60000001b010000270000007468656d652f7468656d652f5f72656c732f7468656d654d616e616765722e786d6c2e72656c73848f4d0ac2301484f78277086f6fd3ba109126dd88d0add40384e4350d363f2451eced0dae2c082e8761be9969bb979dc9136332de3168aa1a083ae995719ac16db8ec8e4052164e89d93b64b060828e6f37ed1567914b284d262452282e3198720e274a939cd08a54f980ae38a38f56e422a3a641c8bbd048f7757da0f19b017cc524bd62107bd5001996509affb3fd381a89672f1f165dfe514173d9850528a2c6cce0239baa4c04ca5bbabac4df000000ffff0300504b01022d0014000600080000002100e9de0fbfff0000001c0200001300000000000000000000000000000000005b436f6e74656e745f54797065735d2e786d6c504b01022d0014000600080000002100a5d6a7e7c0000000360100000b00000000000000000000000000300100005f72656c732f2e72656c73504b01022d00140006000800000021006b799616830000008a0000001c00000000000000000000000000190200007468656d652f7468656d652f7468656d654d616e616765722e786d6c504b01022d0014000600080000002100a91cc9bc48070000022100001600000000000000000000000000d60200007468656d652f7468656d652f7468656d65312e786d6c504b01022d00140006000800000021000dd1909fb60000001b0100002700000000000000000000000000520a00007468656d652f7468656d652f5f72656c732f7468656d654d616e616765722e786d6c2e72656c73504b050600000000050005005d0100004d0b00000000");

pub fn serialize_ext_0896() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_0896_0.serialize());
    buf
}

pub const EXT_089B_0: HexBiffRecord = HexBiffRecord::new(0x089B, "9b080000000000000000000001000000");

pub fn serialize_ext_089b() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_089B_0.serialize());
    buf
}

pub const EXT_088C_0: HexBiffRecord = HexBiffRecord::new(0x088C, "8c080000000000000000000000000000");

pub fn serialize_ext_088c() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_088C_0.serialize());
    buf
}

pub const EXT_00FF_0: HexBiffRecord = HexBiffRecord::new(0x00FF, "0800c82a00000c000000e12b000025010000");

pub fn serialize_ext_00ff() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&EXT_00FF_0.serialize());
    buf
}

// ===== 批量串行化扩展记录（分段） =====

/// Group A: 位于 BOUNDSHEET 之前
pub fn serialize_ext_group_a() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&serialize_ext_087c());
    buf.extend_from_slice(&serialize_ext_087d());
    buf.extend_from_slice(&serialize_ext_0293_0892_interleaved());
    buf.extend_from_slice(&serialize_ext_088e());
    buf.extend_from_slice(&serialize_ext_0160());
    buf
}

/// Group B: 位于 BOUNDSHEET 和 SST 之间
pub fn serialize_ext_group_b() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&serialize_ext_089a());
    buf.extend_from_slice(&serialize_ext_08a3());
    buf.extend_from_slice(&serialize_ext_008c());
    buf.extend_from_slice(&serialize_ext_01c1());
    buf
}

/// Group C: 位于 SST 和 EOF 之间
pub fn serialize_ext_group_c() -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(&serialize_ext_00ff());
    buf.extend_from_slice(&serialize_ext_0863());
    buf.extend_from_slice(&serialize_ext_0896());
    buf.extend_from_slice(&serialize_ext_089b());
    buf.extend_from_slice(&serialize_ext_088c());
    buf
}
