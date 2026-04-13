use super::BiffRecord;

/// FnGroupCountRecord 记录
///
/// 作用：定义内置函数组的数量
///
/// FnGroupCountRecord是Excel BIFF格式中的函数组计数记录（ID: 0x009C），
/// 用于指定Excel内置的用户定义函数组数量。
///
/// ## 参数说明
///
/// - 固定值：0x000E (14个函数组)
#[derive(Debug, Default)]
pub struct FnGroupCountRecord;

impl FnGroupCountRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for FnGroupCountRecord {
    fn id(&self) -> u16 {
        0x009C // FnGroupCount record ID
    }

    fn data(&self) -> Vec<u8> {
        vec![0x0E, 0x00] // Hard-coded 0x000E (14 function groups)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_fn_group_count_record_id() {
        let record = FnGroupCountRecord::new();
        assert_eq!(record.id(), 0x009C);
    }

    #[test]
    fn test_fn_group_count_record_default() {
        let record = FnGroupCountRecord::default();
        assert_eq!(record.id(), 0x009C);
    }

    #[test]
    fn test_fn_group_count_record_data_size() {
        let record = FnGroupCountRecord::new();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_fn_group_count_record_data_value() {
        let record = FnGroupCountRecord::new();
        assert_eq!(&record.data()[..], &0x000Eu16.to_le_bytes());
    }
}
