use super::encode_biff_string_v2;
use super::BiffRecord;

#[derive(Debug)]
pub struct NumberFormatRecord {
    format_index: u16,
    format_string: String,
}

impl NumberFormatRecord {
    pub fn new(format_index: u16, format_string: &str) -> Self {
        NumberFormatRecord {
            format_index,
            format_string: format_string.to_string(),
        }
    }
}

impl BiffRecord for NumberFormatRecord {
    fn id(&self) -> u16 {
        0x041E
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::new();

        // 2 bytes: format index
        buf.extend_from_slice(&self.format_index.to_le_bytes());

        // Format string using upack2 logic
        let encoded_string = encode_biff_string_v2(&self.format_string);
        buf.extend_from_slice(&encoded_string);

        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_number_format_record_id() {
        let record = NumberFormatRecord::new(0, "general");
        assert_eq!(record.id(), 0x041E);
    }

    #[test]
    fn test_number_format_record_data() {
        let record = NumberFormatRecord::new(164, "0.00%");
        let data = record.data();

        // Check format index (2 bytes)
        assert_eq!(&data[0..2], &164u16.to_le_bytes());

        // Check string length (2 bytes) - "0.00%" = 5 chars
        assert_eq!(&data[2..4], &5u16.to_le_bytes());

        // Check flag (1 byte) - "0.00%" is ASCII, so flag = 0
        assert_eq!(data[4], 0x00);
    }

    #[test]
    fn test_number_format_record_data_size() {
        let record = NumberFormatRecord::new(0, "general");
        let data = record.data();

        // "general" is ASCII: 2 (index) + 2 (len) + 1 (flag) + 7 (ASCII) = 12 bytes
        assert_eq!(data.len(), 12);
    }

    #[test]
    fn test_number_format_record_unicode() {
        let record = NumberFormatRecord::new(164, "人民币");
        let data = record.data();

        // Unicode uses UTF-16LE: 2 (index) + 2 (len) + 1 (flag) + 6 (UTF-16 "人民币" = 3 * 2)
        assert_eq!(data.len(), 11);

        // Check flag = 1 for Unicode
        assert_eq!(data[4], 0x01);
    }
}
