// src/biff_records/codepage_record.rs

use super::BiffRecord;

/// CodepageRecord 记录
///
/// 作用：定义工作簿的字符编码
///
/// CodepageRecord是Excel BIFF格式中的代码页记录（ID: 0x0042），用于定义
/// 工作簿中字符串的字符编码方式。
///
/// ## 参数说明
///
/// - 固定值：0x04B0 (1200 = UTF-16)
/// - 此记录强制使用UTF-16编码
#[derive(Debug, Default)]
pub struct CodepageRecord;

impl CodepageRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for CodepageRecord {
    fn id(&self) -> u16 {
        0x0042 // Codepage record ID
    }

    fn data(&self) -> Vec<u8> {
        vec![0xB0, 0x04] // Hard-coded UTF-16 (1200 in little-endian)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_codepage_record_id() {
        let record = CodepageRecord::new();
        assert_eq!(record.id(), 0x0042);
    }

    #[test]
    fn test_codepage_record_default() {
        let record = CodepageRecord::default();
        assert_eq!(record.id(), 0x0042);
    }

    #[test]
    fn test_codepage_record_data_size() {
        let record = CodepageRecord::new();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_codepage_record_data_value() {
        let record = CodepageRecord::new();
        assert_eq!(&record.data()[..], &0x04B0u16.to_le_bytes());
    }
}
