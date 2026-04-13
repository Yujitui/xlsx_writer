use super::BiffRecord;

/// WindowProtectRecord 记录（工作簿窗口保护）
///
/// 作用：控制工作簿窗口的保护状态
///
/// WindowProtectRecord是Excel BIFF格式中的窗口保护记录（ID: 0x0019），
/// 用于保护工作簿窗口的布局，防止用户改变窗口大小或位置。
///
/// ## 参数说明
///
/// - `is_protected`: 保护标志
///   - false = 未保护（默认）
///   - true = 保护窗口布局
#[derive(Debug)]
pub struct WindowProtectRecord {
    is_protected: bool,
}

impl WindowProtectRecord {
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

impl Default for WindowProtectRecord {
    fn default() -> Self {
        Self::unprotected()
    }
}

impl BiffRecord for WindowProtectRecord {
    fn id(&self) -> u16 {
        0x0019 // WindowProtect record ID
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
    fn test_window_protect_record_id() {
        let record = WindowProtectRecord::new(false);
        assert_eq!(record.id(), 0x0019);
    }

    #[test]
    fn test_window_protect_record_default() {
        let record = WindowProtectRecord::default();
        assert_eq!(record.is_protected, false);
    }

    #[test]
    fn test_window_protect_record_data_size() {
        let record = WindowProtectRecord::new(false);
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_window_protect_record_protected() {
        let record = WindowProtectRecord::protected();
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }
}
