use super::BiffRecord;

/// ObjectProtectRecord 记录（工作簿对象保护）
///
/// 作用：控制工作簿中对象的保护状态
///
/// ObjectProtectRecord是Excel BIFF格式中的对象保护记录（ID: 0x0063），
/// 用于定义工作簿中非单元格对象（如图表、图形、文本框等）的保护状态。
///
/// ## 参数说明
///
/// - `is_protected`: 保护标志
///   - false = 未保护（默认）
///   - true = 保护
#[derive(Debug)]
pub struct ObjectProtectRecord {
    is_protected: bool,
}

impl ObjectProtectRecord {
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

impl Default for ObjectProtectRecord {
    fn default() -> Self {
        Self::unprotected()
    }
}

impl BiffRecord for ObjectProtectRecord {
    fn id(&self) -> u16 {
        0x0063 // ObjectProtect record ID
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
    fn test_object_protect_record_id() {
        let record = ObjectProtectRecord::new(false);
        assert_eq!(record.id(), 0x0063);
    }

    #[test]
    fn test_object_protect_record_default() {
        let record = ObjectProtectRecord::default();
        assert_eq!(record.is_protected, false);
    }

    #[test]
    fn test_object_protect_record_data_size() {
        let record = ObjectProtectRecord::new(false);
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_object_protect_record_protected() {
        let record = ObjectProtectRecord::protected();
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }

    #[test]
    fn test_object_protect_record_unprotected() {
        let record = ObjectProtectRecord::unprotected();
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }
}
