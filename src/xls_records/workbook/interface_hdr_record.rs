use super::BiffRecord;

/// InterfaceHdrRecord 记录
///
/// 作用：标记HTML UI界面区域的开始
///
/// InterfaceHdrRecord是Excel BIFF格式中的界面头部记录（ID: 0x00E1），
/// 用于标记HTML UI（用户界面）格式数据区域的开始，并指定代码页。
///
/// ## 参数说明
///
/// - 固定值：0x04B0 (1200 = UTF-16代码页)
#[derive(Debug)]
pub struct InterfaceHdrRecord;

impl InterfaceHdrRecord {
    pub fn new() -> Self {
        Self
    }
}

impl Default for InterfaceHdrRecord {
    fn default() -> Self {
        Self
    }
}

impl BiffRecord for InterfaceHdrRecord {
    fn id(&self) -> u16 {
        0x00E1 // Interface Header record ID
    }

    fn data(&self) -> Vec<u8> {
        vec![0xB0, 0x04] // Corrected: 0x04B0 (1200 decimal)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_interface_hdr_record_id() {
        let record = InterfaceHdrRecord::new();
        assert_eq!(record.id(), 0x00E1);
    }

    #[test]
    fn test_interface_hdr_record_default() {
        let record = InterfaceHdrRecord::default();
        assert_eq!(record.id(), 0x00E1);
    }

    #[test]
    fn test_interface_hdr_record_data_size() {
        let record = InterfaceHdrRecord::new();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_interface_hdr_record_data_value() {
        let record = InterfaceHdrRecord::new();
        assert_eq!(&record.data()[..], &0x04B0u16.to_le_bytes());
    }
}
