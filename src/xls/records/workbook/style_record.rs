use super::BiffRecord;

/// StyleRecord 记录
///
/// 作用：定义用户定义的样式
///
/// StyleRecord是Excel BIFF格式中的样式记录（ID: 0x0293），用于定义
/// 用户创建的命名样式。此记录已较少使用。
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
