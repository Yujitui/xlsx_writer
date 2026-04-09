use crate::xls::records::BiffRecord;

/// ProtectRecord 记录（工作表保护）
///
/// 作用：存储工作表的保护状态
///
/// ProtectRecord是Excel BIFF格式中的工作表保护记录（ID: 0x0012），用于定义
/// 工作表是否被保护。当保护启用时，用户无法修改受保护的单元格和结构。
///
/// ## 参数说明
///
/// - `value`: 保护标志（0=未保护, 1=已保护）
/// - 默认值为 0（未保护）
#[derive(Debug)]
pub struct ProtectRecord {
    value: u16,
}

impl ProtectRecord {
    pub fn new(value: u16) -> Self {
        ProtectRecord { value }
    }
}

impl Default for ProtectRecord {
    fn default() -> Self {
        ProtectRecord { value: 0 }
    }
}

impl BiffRecord for ProtectRecord {
    fn id(&self) -> u16 {
        0x0012
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
    fn test_protect_record_id() {
        let record = ProtectRecord::default();
        assert_eq!(record.id(), 0x0012);
    }

    #[test]
    fn test_protect_record_default() {
        let record = ProtectRecord::default();
        assert_eq!(record.value, 0);
    }

    #[test]
    fn test_protect_record_data_size() {
        let record = ProtectRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
