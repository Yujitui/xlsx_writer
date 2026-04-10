use super::encode_biff_string_v1;
use super::BiffRecord;

/// BoundSheetRecord 记录
///
/// 作用：定义工作簿中工作表的信息
///
/// BoundSheetRecord是Excel BIFF格式中的工作表信息记录（ID: 0x0085），用于
/// 定义每个工作表的名称、类型、可见性和在文件流中的位置。
///
/// ## 参数说明
///
/// - `stream_pos`: 工作表数据在文件流中的位置偏移
/// - `visibility`: 可见性（见SheetVisibility枚举）
/// - `sheet_type`: 工作表类型（见SheetType枚举）
/// - `sheet_name`: 工作表名称
#[derive(Debug)]
pub struct BoundSheetRecord {
    stream_pos: u32,
    visibility: u8,
    sheet_type: u8,
    sheet_name: String,
}

#[repr(u8)]
#[derive(Debug, Clone, Copy)]
pub enum SheetVisibility {
    Visible = 0x00,
    Hidden = 0x01,
    StrongHidden = 0x02,
}

#[repr(u8)]
#[derive(Debug, Clone, Copy)]
pub enum SheetType {
    Worksheet = 0x00,
    Chart = 0x02,
    VbaModule = 0x06,
}

impl BoundSheetRecord {
    pub fn new(
        stream_pos: u32,
        visibility: SheetVisibility,
        sheet_type: SheetType,
        sheet_name: &str,
    ) -> Self {
        BoundSheetRecord {
            stream_pos,
            visibility: visibility as u8,
            sheet_type: sheet_type as u8,
            sheet_name: sheet_name.to_string(),
        }
    }

    pub fn new_pending(sheet_name: &str) -> Self {
        BoundSheetRecord {
            stream_pos: 0,
            visibility: SheetVisibility::Visible as u8,
            sheet_type: SheetType::Worksheet as u8,
            sheet_name: sheet_name.to_string(),
        }
    }

    pub fn set_offset(&mut self, offset: u32) {
        self.stream_pos = offset;
    }
}

impl BiffRecord for BoundSheetRecord {
    fn id(&self) -> u16 {
        0x0085
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::new();

        // 4 bytes: absolute stream position
        buf.extend_from_slice(&self.stream_pos.to_le_bytes());

        // 1 byte: visibility
        buf.push(self.visibility);

        // 1 byte: sheet type
        buf.push(self.sheet_type);

        // Sheet name using upack1 logic
        let encoded_name = encode_biff_string_v1(&self.sheet_name);
        buf.extend_from_slice(&encoded_name);

        buf
    }
}

// ============================================================================
// ParsableRecord implementation for reading
// ============================================================================

use crate::xls::records::{ParsableRecord, ParseState};
use crate::xls::XlsError;

impl ParsableRecord for BoundSheetRecord {
    const RECORD_ID: u16 = 0x0085;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 6 {
            return Err(XlsError::InvalidFormat(format!(
                "BoundSheetRecord data too short: {} bytes",
                data.len()
            )));
        }

        let stream_pos = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
        let visibility = data[4];
        let sheet_type = data[5];

        // 解析工作表名称（使用 upack1 格式：1字节长度 + 1字节标志 + 数据）
        let mut offset = 6;

        if offset >= data.len() {
            return Err(XlsError::InvalidFormat(
                "BoundSheetRecord name length missing".to_string(),
            ));
        }

        let name_len = data[offset] as usize;
        offset += 1;

        if offset >= data.len() {
            return Err(XlsError::InvalidFormat(
                "BoundSheetRecord name flag missing".to_string(),
            ));
        }

        let flag = data[offset];
        let is_utf16 = (flag & 0x01) != 0;
        offset += 1;

        let sheet_name = if is_utf16 {
            // UTF-16LE 解码
            if offset + name_len * 2 > data.len() {
                return Err(XlsError::InvalidFormat(
                    "BoundSheetRecord name data incomplete".to_string(),
                ));
            }
            let utf16_data = &data[offset..offset + name_len * 2];
            let u16_vec: Vec<u16> = utf16_data
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect();
            String::from_utf16(&u16_vec)
                .unwrap_or_else(|_| String::from_utf8_lossy(utf16_data).to_string())
        } else {
            // ASCII/Latin-1 解码
            if offset + name_len > data.len() {
                return Err(XlsError::InvalidFormat(
                    "BoundSheetRecord name data incomplete".to_string(),
                ));
            }
            String::from_utf8_lossy(&data[offset..offset + name_len]).to_string()
        };

        Ok(BoundSheetRecord {
            stream_pos,
            visibility,
            sheet_type,
            sheet_name,
        })
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        // 将工作表名称添加到列表
        state.sheet_names.push(self.sheet_name.clone());
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_boundsheet_record_id() {
        let record =
            BoundSheetRecord::new(0, SheetVisibility::Visible, SheetType::Worksheet, "Sheet1");
        assert_eq!(record.id(), 0x0085);
    }

    #[test]
    fn test_boundsheet_record_basic() {
        let record = BoundSheetRecord::new(
            256,
            SheetVisibility::Visible,
            SheetType::Worksheet,
            "Sheet1",
        );
        let data = record.data();

        // Check stream position (4 bytes)
        assert_eq!(&data[0..4], &256u32.to_le_bytes());

        // Check visibility (1 byte)
        assert_eq!(data[4], 0x00);

        // Check sheet type (1 byte)
        assert_eq!(data[5], 0x00);

        // Check name length (1 byte)
        assert_eq!(data[6], 6); // "Sheet1" = 6 chars

        // Check flag (1 byte) - ASCII uses flag 0
        assert_eq!(data[7], 0x00);
    }

    #[test]
    fn test_boundsheet_record_unicode() {
        let record =
            BoundSheetRecord::new(0, SheetVisibility::Visible, SheetType::Worksheet, "表格");
        let data = record.data();

        // Check visibility (1 byte)
        assert_eq!(data[4], 0x00);

        // Check sheet type (1 byte)
        assert_eq!(data[5], 0x00);

        // Check name length (2 for UTF-16)
        assert_eq!(data[6], 2);

        // Check flag (1 byte) - Unicode uses flag 1
        assert_eq!(data[7], 0x01);
    }
}
