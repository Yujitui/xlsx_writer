use crate::xls::records::BiffRecord;

/// WindowProtectRecord 记录（工作表窗口保护）
///
/// 作用：控制工作表窗口的保护状态
///
/// WindowProtectRecord是Excel BIFF格式中的窗口保护记录（ID: 0x0019），
/// 用于保护工作表窗口的显示状态，防止用户改变冻结窗格等设置。
///
/// ## 参数说明
///
/// - `value`: 保护标志
///   - 0 = 未保护（默认）
///   - 1 = 已保护
#[derive(Debug)]
pub struct WindowProtectRecord {
    value: u16,
}

impl WindowProtectRecord {
    pub fn new(value: u16) -> Self {
        WindowProtectRecord { value }
    }
}

impl Default for WindowProtectRecord {
    fn default() -> Self {
        WindowProtectRecord { value: 0 }
    }
}

impl BiffRecord for WindowProtectRecord {
    fn id(&self) -> u16 {
        0x0019
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
    fn test_window_protect_record_id() {
        let record = WindowProtectRecord::default();
        assert_eq!(record.id(), 0x0019);
    }

    #[test]
    fn test_window_protect_record_default() {
        let record = WindowProtectRecord::default();
        assert_eq!(record.value, 0);
    }

    #[test]
    fn test_window_protect_record_data_size() {
        let record = WindowProtectRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
