use crate::xls_records::BiffRecord;

/// WSBoolRecord 记录
///
/// 作用：存储工作表的布尔选项设置
///
/// WSBoolRecord是Excel BIFF格式中的工作表布尔选项记录（ID: 0x0081），用于
/// 定义工作表的各种显示和操作选项。
///
/// ## 参数说明
///
/// - `options`: 选项标志
///   - bit 0 (0x0001): 显示图表
///   - bit 1 (0x0002): 显示对话框
///   - bit 2 (0x0004): 冻结窗格
///   - bit 10 (0x0400): 显示零值
///   - bit 12 (0x1000): 自动打印标题
/// - 默认值为 0x0C01
#[derive(Debug)]
pub struct WSBoolRecord {
    options: u16,
}

impl WSBoolRecord {
    pub fn new(options: u16) -> Self {
        WSBoolRecord { options }
    }
}

impl Default for WSBoolRecord {
    fn default() -> Self {
        WSBoolRecord { options: 0x0C01 }
    }
}

impl BiffRecord for WSBoolRecord {
    fn id(&self) -> u16 {
        0x0081
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(2);
        buf.extend_from_slice(&self.options.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_wsbool_record_id() {
        let record = WSBoolRecord::default();
        assert_eq!(record.id(), 0x0081);
    }

    #[test]
    fn test_wsbool_record_default() {
        let record = WSBoolRecord::default();
        assert_eq!(record.options, 0x0C01);
    }

    #[test]
    fn test_wsbool_record_data_size() {
        let record = WSBoolRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
