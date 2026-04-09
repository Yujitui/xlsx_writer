use super::BiffRecord;

/// InterfaceEndRecord 记录
///
/// 作用：标记HTML UI界面区域的结束
///
/// InterfaceEndRecord是Excel BIFF格式中的界面结束记录（ID: 0x00E2），
/// 用于标记HTML UI（用户界面）格式数据区域的结束。无数据部分。
#[derive(Debug, Default)]
pub struct InterfaceEndRecord;

impl InterfaceEndRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for InterfaceEndRecord {
    fn id(&self) -> u16 {
        0x00E2 // Interface End record ID
    }

    fn data(&self) -> Vec<u8> {
        Vec::new() // Empty data
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_interface_end_record_id() {
        let record = InterfaceEndRecord::new();
        assert_eq!(record.id(), 0x00E2);
    }

    #[test]
    fn test_interface_end_record_default() {
        let record = InterfaceEndRecord::default();
        assert_eq!(record.id(), 0x00E2);
    }

    #[test]
    fn test_interface_end_record_data_size() {
        let record = InterfaceEndRecord::new();
        assert_eq!(record.data().len(), 0);
    }
}
