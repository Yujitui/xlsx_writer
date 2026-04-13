use crate::xls_records::BiffRecord;

/// HCenterRecord 记录
///
/// 作用：控制打印页面水平居中
///
/// HCenterRecord是Excel BIFF格式中的水平居中记录（ID: 0x0083），用于定义
/// 打印时页面内容是否水平居中。
///
/// ## 参数说明
///
/// - `value`: 水平居中标志
///   - 0 = 不居中
///   - 1 = 水平居中（默认）
#[derive(Debug)]
pub struct HCenterRecord {
    value: u16,
}

impl HCenterRecord {
    pub fn new(value: u16) -> Self {
        HCenterRecord { value }
    }
}

impl Default for HCenterRecord {
    fn default() -> Self {
        HCenterRecord { value: 1 }
    }
}

impl BiffRecord for HCenterRecord {
    fn id(&self) -> u16 {
        0x0083
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(2);
        buf.extend_from_slice(&self.value.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hcenter_record_id() {
        let record = HCenterRecord::default();
        assert_eq!(record.id(), 0x0083);
    }

    #[test]
    fn test_hcenter_record_default() {
        let record = HCenterRecord::default();
        assert_eq!(record.value, 1);
    }

    #[test]
    fn test_hcenter_record_data_size() {
        let record = HCenterRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
