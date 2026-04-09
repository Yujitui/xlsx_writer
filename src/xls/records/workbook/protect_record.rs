use super::BiffRecord;

/// ProtectRecord 记录（工作簿保护）
///
/// 作用：控制工作簿的结构保护
///
/// ProtectRecord是Excel BIFF格式中的工作簿保护记录（ID: 0x0012），用于
/// 保护工作簿的结构，防止用户添加、删除或重命名工作表。
///
/// ## 参数说明
///
/// - `is_protected`: 保护标志
///   - false = 未保护（默认）
///   - true = 保护工作簿结构
#[derive(Debug)]
pub struct ProtectRecord {
    is_protected: bool,
}

impl ProtectRecord {
    pub fn new(is_protected: bool) -> Self {
        Self { is_protected }
    }

    pub fn protected() -> Self {
        Self::new(true)
    }

    pub fn unprotected() -> Self {
        Self::new(false)
    }
}

impl Default for ProtectRecord {
    fn default() -> Self {
        Self::unprotected()
    }
}

impl BiffRecord for ProtectRecord {
    fn id(&self) -> u16 {
        0x0012 // Protect record ID
    }

    fn data(&self) -> Vec<u8> {
        if self.is_protected {
            vec![0x01, 0x00] // 1 = Protected
        } else {
            vec![0x00, 0x00] // 0 = Not protected
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_protect_record_id() {
        let record = ProtectRecord::new(false);
        assert_eq!(record.id(), 0x0012);
    }

    #[test]
    fn test_protect_record_default() {
        let record = ProtectRecord::default();
        assert_eq!(record.is_protected, false);
    }

    #[test]
    fn test_protect_record_data_size() {
        let record = ProtectRecord::new(false);
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_protect_record_protected() {
        let record = ProtectRecord::protected();
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }

    #[test]
    fn test_protect_record_unprotected() {
        let record = ProtectRecord::unprotected();
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }
}
