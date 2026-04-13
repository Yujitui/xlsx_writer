use crate::xls_records::BiffRecord;

/// VCenterRecord 记录
///
/// 作用：控制打印页面垂直居中
///
/// VCenterRecord是Excel BIFF格式中的垂直居中记录（ID: 0x0084），用于定义
/// 打印时页面内容是否垂直居中。
///
/// ## 参数说明
///
/// - `value`: 垂直居中标志
///   - 0 = 不居中（默认）
///   - 1 = 垂直居中
#[derive(Debug)]
pub struct VCenterRecord {
    value: u16,
}

impl VCenterRecord {
    pub fn new(value: u16) -> Self {
        VCenterRecord { value }
    }
}

impl Default for VCenterRecord {
    fn default() -> Self {
        VCenterRecord { value: 0 }
    }
}

impl BiffRecord for VCenterRecord {
    fn id(&self) -> u16 {
        0x0084
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
    fn test_vcenter_record_id() {
        let record = VCenterRecord::default();
        assert_eq!(record.id(), 0x0084);
    }

    #[test]
    fn test_vcenter_record_default() {
        let record = VCenterRecord::default();
        assert_eq!(record.value, 0);
    }

    #[test]
    fn test_vcenter_record_data_size() {
        let record = VCenterRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
