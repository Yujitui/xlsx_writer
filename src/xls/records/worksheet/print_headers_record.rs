use crate::xls::records::BiffRecord;

/// PrintHeadersRecord 记录
///
/// 作用：控制打印时是否包含行列标题
///
/// PrintHeadersRecord是Excel BIFF格式中的打印标题记录（ID: 0x002A），
/// 用于定义打印工作表时是否包含行标题（1,2,3...）和列标题（A,B,C...）。
///
/// ## 参数说明
///
/// - `value`: 打印标题标志
///   - 0 = 不打印标题（默认）
///   - 1 = 打印标题
#[derive(Debug)]
pub struct PrintHeadersRecord {
    value: u16,
}

impl PrintHeadersRecord {
    pub fn new(value: u16) -> Self {
        PrintHeadersRecord { value }
    }
}

impl Default for PrintHeadersRecord {
    fn default() -> Self {
        PrintHeadersRecord { value: 0 }
    }
}

impl BiffRecord for PrintHeadersRecord {
    fn id(&self) -> u16 {
        0x002A
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
    fn test_print_headers_record_id() {
        let record = PrintHeadersRecord::default();
        assert_eq!(record.id(), 0x002A);
    }

    #[test]
    fn test_print_headers_record_default() {
        let record = PrintHeadersRecord::default();
        assert_eq!(record.value, 0);
    }

    #[test]
    fn test_print_headers_record_data_size() {
        let record = PrintHeadersRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
