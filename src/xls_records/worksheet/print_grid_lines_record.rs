use crate::xls_records::BiffRecord;

/// PrintGridLinesRecord 记录
///
/// 作用：控制打印时是否包含网格线
///
/// PrintGridLinesRecord是Excel BIFF格式中的打印网格线记录（ID: 0x002B），
/// 用于定义打印工作表时是否包含网格线。
///
/// ## 参数说明
///
/// - `value`: 打印网格线标志
///   - 0 = 不打印网格线（默认）
///   - 1 = 打印网格线
#[derive(Debug)]
pub struct PrintGridLinesRecord {
    value: u16,
}

impl PrintGridLinesRecord {
    pub fn new(value: u16) -> Self {
        PrintGridLinesRecord { value }
    }
}

impl Default for PrintGridLinesRecord {
    fn default() -> Self {
        PrintGridLinesRecord { value: 0 }
    }
}

impl BiffRecord for PrintGridLinesRecord {
    fn id(&self) -> u16 {
        0x002B
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
    fn test_print_grid_lines_record_id() {
        let record = PrintGridLinesRecord::default();
        assert_eq!(record.id(), 0x002B);
    }

    #[test]
    fn test_print_grid_lines_record_default() {
        let record = PrintGridLinesRecord::default();
        assert_eq!(record.value, 0);
    }

    #[test]
    fn test_print_grid_lines_record_data_size() {
        let record = PrintGridLinesRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
