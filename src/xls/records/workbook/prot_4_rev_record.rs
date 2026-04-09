use super::BiffRecord;

/// Prot4RevRecord 记录
///
/// 作用：控制工作簿的修订保护状态
///
/// Prot4RevRecord是Excel BIFF格式中的修订保护记录（ID: 0x01AF），用于
/// 启用或禁用工作簿的修订保护功能。修订保护允许多人协作跟踪更改。
///
/// ## 参数说明
///
/// - 固定值：0x0000（默认，未启用修订保护）
#[derive(Debug, Clone, PartialEq)]
pub struct Prot4RevRecord;

impl Prot4RevRecord {
    pub fn new() -> Self {
        Self
    }
}

impl Default for Prot4RevRecord {
    fn default() -> Self {
        Self
    }
}

impl BiffRecord for Prot4RevRecord {
    fn id(&self) -> u16 {
        0x01AF
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_prot_4_rev_record_id() {
        let record = Prot4RevRecord::new();
        assert_eq!(record.id(), 0x01AF);
    }

    #[test]
    fn test_prot_4_rev_record_default() {
        let record = Prot4RevRecord::default();
        assert_eq!(record.id(), 0x01AF);
    }

    #[test]
    fn test_prot_4_rev_record_data_size() {
        let record = Prot4RevRecord::new();
        assert_eq!(record.data().len(), 2);
    }
}
