use crate::xls_records::BiffRecord;

/// ObjectProtectRecord 记录（工作表对象保护）
///
/// 作用：存储工作表中对象的保护状态
///
/// ObjectProtectRecord是Excel BIFF格式中的对象保护记录（ID: 0x0063），
/// 用于保护工作表中的非单元格对象（如图表、图形等）。
///
/// ## 参数说明
///
/// - `value`: 保护标志
///   - 0 = 未保护（默认）
///   - 1 = 已保护
#[derive(Debug)]
pub struct ObjectProtectRecord {
    value: u16,
}

impl ObjectProtectRecord {
    pub fn new(value: u16) -> Self {
        ObjectProtectRecord { value }
    }
}

impl Default for ObjectProtectRecord {
    fn default() -> Self {
        ObjectProtectRecord { value: 0 }
    }
}

impl BiffRecord for ObjectProtectRecord {
    fn id(&self) -> u16 {
        0x0063
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
    fn test_object_protect_record_id() {
        let record = ObjectProtectRecord::default();
        assert_eq!(record.id(), 0x0063);
    }

    #[test]
    fn test_object_protect_record_default() {
        let record = ObjectProtectRecord::default();
        assert_eq!(record.value, 0);
    }

    #[test]
    fn test_object_protect_record_data_size() {
        let record = ObjectProtectRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
