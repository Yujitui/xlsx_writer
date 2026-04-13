use super::BiffRecord;

/// DSFRecord 记录
///
/// 作用：表示双文档流标志
///
/// DSFRecord是Excel BIFF格式中的双文档流标志记录（ID: 0x0161）。
/// 固定值0x0000表示不使用双文档流（正常模式）。
///
/// ## 参数说明
///
/// - 固定值：0x0000
#[derive(Debug, Default)]
pub struct DSFRecord;

impl DSFRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for DSFRecord {
    fn id(&self) -> u16 {
        0x0161 // DSF record ID
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00] // Hard-coded 0x0000
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_dsf_record_id() {
        let record = DSFRecord::new();
        assert_eq!(record.id(), 0x0161);
    }

    #[test]
    fn test_dsf_record_default() {
        let record = DSFRecord::default();
        assert_eq!(record.id(), 0x0161);
    }

    #[test]
    fn test_dsf_record_data_size() {
        let record = DSFRecord::new();
        assert_eq!(record.data().len(), 2);
    }
}
