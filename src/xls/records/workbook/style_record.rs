use super::BiffRecord;

#[derive(Debug, Default)]
pub struct StyleRecord;

impl StyleRecord {
    pub fn new() -> Self {
        StyleRecord
    }
}

impl BiffRecord for StyleRecord {
    fn id(&self) -> u16 {
        0x0293
    }

    fn data(&self) -> Vec<u8> {
        // Default: header = 0x8000 (user-defined, index = 0xFFF)
        // Built-in flag = 0 (user-defined)
        let mut buf = Vec::with_capacity(3);
        buf.extend_from_slice(&0x8000u16.to_le_bytes());
        buf.push(0x00); // Built-in identifier
        buf.push(0xFF); // Level (0xFF for non-built-in)
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_style_record_id() {
        let record = StyleRecord::new();
        assert_eq!(record.id(), 0x0293);
    }

    #[test]
    fn test_style_record_data() {
        let record = StyleRecord::new();
        let data = record.data();

        // Check header (2 bytes)
        assert_eq!(&data[0..2], &0x8000u16.to_le_bytes());

        // Check built-in identifier (1 byte)
        assert_eq!(data[2], 0x00);

        // Check level (1 byte)
        assert_eq!(data[3], 0xFF);
    }

    #[test]
    fn test_style_record_data_size() {
        let record = StyleRecord::new();
        assert_eq!(record.data().len(), 4);
    }
}
