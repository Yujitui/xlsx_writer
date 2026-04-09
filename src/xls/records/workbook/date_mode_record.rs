use super::BiffRecord;

/// DateModeRecord 记录
///
/// 作用：定义工作簿的日期系统
///
/// DateModeRecord是Excel BIFF格式中的日期模式记录（ID: 0x0022），用于定义
/// 工作簿中日期值所使用的时间系统（1900或1904）。
///
/// ## 参数说明
///
/// - `from_1904`: 日期系统选择
///   - false = 1900日期系统（默认，Excel默认）
///   - true = 1904日期系统（Macintosh Excel早期版本）
#[derive(Debug)]
pub struct DateModeRecord {
    from_1904: bool,
}

impl DateModeRecord {
    pub fn new(from_1904: bool) -> Self {
        DateModeRecord { from_1904 }
    }
}

impl Default for DateModeRecord {
    fn default() -> Self {
        Self::new(false)
    }
}

impl BiffRecord for DateModeRecord {
    fn id(&self) -> u16 {
        0x0022
    }

    fn data(&self) -> Vec<u8> {
        let value: u16 = if self.from_1904 { 1 } else { 0 };
        value.to_le_bytes().to_vec()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_date_mode_record_id() {
        let record = DateModeRecord::new(false);
        assert_eq!(record.id(), 0x0022);
    }

    #[test]
    fn test_date_mode_record_data_size() {
        let record = DateModeRecord::new(false);
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_date_mode_from_1904_false() {
        let record = DateModeRecord::new(false);
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }

    #[test]
    fn test_date_mode_from_1904_true() {
        let record = DateModeRecord::new(true);
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }
}
