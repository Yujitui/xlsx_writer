use super::BiffRecord;

/// Prot4RevPassRecord 记录
///
/// 作用：存储工作簿修订保护密码
///
/// Prot4RevPassRecord是Excel BIFF格式中的修订保护密码记录（ID: 0x01BC），
/// 用于存储工作簿修订保护的密码哈希。固定值0x0000表示无密码。
///
/// ## 参数说明
///
/// - 固定值：0x0000（无密码）
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Prot4RevPassRecord;

impl Prot4RevPassRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for Prot4RevPassRecord {
    fn id(&self) -> u16 {
        0x01BC
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_prot_4_rev_pass_record_id() {
        let record = Prot4RevPassRecord::new();
        assert_eq!(record.id(), 0x01BC);
    }

    #[test]
    fn test_prot_4_rev_pass_record_data_size() {
        let record = Prot4RevPassRecord::new();
        assert_eq!(record.data().len(), 2);
    }
}
