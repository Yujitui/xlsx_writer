use super::encode_biff_string_v1;
use super::BiffRecord;

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
