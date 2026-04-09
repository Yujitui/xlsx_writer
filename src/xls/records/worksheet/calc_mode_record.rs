use crate::xls::records::BiffRecord;

/// CalcModeRecord 记录
///
/// 作用：控制工作表的计算模式
///
/// CalcModeRecord是Excel BIFF格式中的计算模式记录（ID: 0x000D），用于定义
/// 工作表中公式的计算方式（自动计算或手动计算）。
///
/// ## 参数说明
///
/// - `calc_mode`: 计算模式
///   - 1 = 自动计算
///   - 2 = 按需计算（手动）
/// - 默认值为 1（自动计算）
#[derive(Debug)]
pub struct CalcModeRecord {
    calc_mode: i16,
}

impl CalcModeRecord {
    pub fn new(calc_mode: i16) -> Self {
        CalcModeRecord { calc_mode }
    }
}

impl Default for CalcModeRecord {
    fn default() -> Self {
        CalcModeRecord { calc_mode: 1 }
    }
}

impl BiffRecord for CalcModeRecord {
    fn id(&self) -> u16 {
        0x000D
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(2);
        buf.extend_from_slice(&self.calc_mode.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_calc_mode_record_id() {
        let record = CalcModeRecord::default();
        assert_eq!(record.id(), 0x000D);
    }

    #[test]
    fn test_calc_mode_record_default() {
        let record = CalcModeRecord::default();
        assert_eq!(record.calc_mode, 1);
    }

    #[test]
    fn test_calc_mode_record_data_size() {
        let record = CalcModeRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
