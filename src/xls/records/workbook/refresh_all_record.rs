use super::BiffRecord;

/// RefreshAllRecord 记录
///
/// 作用：控制是否刷新所有数据透视表
///
/// RefreshAllRecord是Excel BIFF格式中的刷新全部记录（ID: 0x01B7），用于
/// 指示Excel在打开文件时是否自动刷新所有数据透视表和外部数据连接。
///
/// ## 参数说明
///
/// - 固定值：0x0000（不自动刷新）
#[derive(Debug)]
pub struct RefreshAllRecord;

impl RefreshAllRecord {
    pub fn new() -> Self {
        RefreshAllRecord
    }
}

impl BiffRecord for RefreshAllRecord {
    fn id(&self) -> u16 {
        0x01B7
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_refresh_all_record_id() {
        let record = RefreshAllRecord::new();
        assert_eq!(record.id(), 0x01B7);
    }

    #[test]
    fn test_refresh_all_record_data_size() {
        let record = RefreshAllRecord::new();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_refresh_all_record_data() {
        let record = RefreshAllRecord::new();
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }
}
