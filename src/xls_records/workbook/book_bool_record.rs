use super::BiffRecord;

/// BookBoolRecord 记录
///
/// 作用：存储工作簿级别的布尔选项
///
/// BookBoolRecord是Excel BIFF格式中的工作簿布尔选项记录（ID: 0x00DA），
/// 用于存储工作簿级别的某些全局选项。数据固定为0x0000。
#[derive(Debug, Default)]
pub struct BookBoolRecord;

impl BookBoolRecord {
    pub fn new() -> Self {
        BookBoolRecord
    }
}

impl BiffRecord for BookBoolRecord {
    fn id(&self) -> u16 {
        0x00DA
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_book_bool_record_id() {
        let record = BookBoolRecord::new();
        assert_eq!(record.id(), 0x00DA);
    }

    #[test]
    fn test_book_bool_record_data_size() {
        let record = BookBoolRecord::new();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_book_bool_record_data() {
        let record = BookBoolRecord::new();
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }
}
