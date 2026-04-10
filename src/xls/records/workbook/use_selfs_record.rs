use super::BiffRecord;

/// UseSelfsRecord 记录
///
/// 作用：指定使用自身引用样式
///
/// UseSelfsRecord是Excel BIFF格式中的使用自身引用记录（ID: 0x0160），
/// 指示公式使用自身引用（R1C1样式）还是相对自身引用。此记录已弃用。
///
/// ## 参数说明
///
/// - 固定值：0x0001（使用!操作符解析自身引用）
#[derive(Debug, Default)]
pub struct UseSelfsRecord;

impl UseSelfsRecord {
    pub fn new() -> Self {
        UseSelfsRecord
    }
}

impl BiffRecord for UseSelfsRecord {
    fn id(&self) -> u16 {
        0x0160
    }

    fn data(&self) -> Vec<u8> {
        vec![0x01, 0x00]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_use_selfs_record_id() {
        let record = UseSelfsRecord::new();
        assert_eq!(record.id(), 0x0160);
    }

    #[test]
    fn test_use_selfs_record_data_size() {
        let record = UseSelfsRecord::new();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_use_selfs_record_data() {
        let record = UseSelfsRecord::new();
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }
}
