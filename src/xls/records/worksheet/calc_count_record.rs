use crate::xls::records::BiffRecord;

/// CalcCountRecord 记录
///
/// 作用：设置工作表的迭代计算次数
///
/// CalcModeRecord是Excel BIFF格式中的迭代次数记录（ID: 0x000C），用于定义
/// 当工作表启用迭代计算时的最大迭代次数。
///
/// ## 参数说明
///
/// - `calc_count`: 最大迭代次数
///   - 默认值为 0x0064 = 100次
///   - 设置为0表示禁用迭代计算
#[derive(Debug)]
pub struct CalcCountRecord {
    calc_count: u16,
}

impl CalcCountRecord {
    pub fn new(calc_count: u16) -> Self {
        CalcCountRecord { calc_count }
    }
}

impl Default for CalcCountRecord {
    fn default() -> Self {
        CalcCountRecord { calc_count: 0x0064 }
    }
}

impl BiffRecord for CalcCountRecord {
    fn id(&self) -> u16 {
        0x000C
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(2);
        buf.extend_from_slice(&self.calc_count.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_calc_count_record_id() {
        let record = CalcCountRecord::default();
        assert_eq!(record.id(), 0x000C);
    }

    #[test]
    fn test_calc_count_record_default() {
        let record = CalcCountRecord::default();
        assert_eq!(record.calc_count, 0x0064);
    }
}
