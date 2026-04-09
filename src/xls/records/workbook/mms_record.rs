use super::BiffRecord;

/// MMSRecord 记录
///
/// 作用：存储菜单和工具栏数量信息
///
/// MMSRecord是Excel BIFF格式中的MMS记录（ID: 0x00C1），用于存储自定义
/// 菜单和工具栏的数量信息。此记录已弃用，固定值为0。
///
/// ## 参数说明
///
/// - 固定值：0x0000
#[derive(Debug, Default)]
pub struct MMSRecord;

impl MMSRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for MMSRecord {
    fn id(&self) -> u16 {
        0x00C1 // MMS record ID
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00] // 2 bytes: 0x0000
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_mms_record_id() {
        let record = MMSRecord::new();
        assert_eq!(record.id(), 0x00C1);
    }

    #[test]
    fn test_mms_record_default() {
        let record = MMSRecord::default();
        assert_eq!(record.id(), 0x00C1);
    }

    #[test]
    fn test_mms_record_data_size() {
        let record = MMSRecord::new();
        assert_eq!(record.data().len(), 2);
    }
}
