use crate::xls_records::BiffRecord;

/// IterationRecord 记录
///
/// 作用：控制工作表的迭代计算开关
///
/// IterationRecord是Excel BIFF格式中的迭代开关记录（ID: 0x011），用于启用或
/// 禁用工作表中的迭代计算。迭代计算用于解决循环引用问题。
///
/// ## 参数说明
///
/// - `iterations_on`: 迭代开关
///   - 0 = 禁用迭代计算（默认）
///   - 1 = 启用迭代计算
#[derive(Debug)]
pub struct IterationRecord {
    iterations_on: u16,
}

impl IterationRecord {
    pub fn new(iterations_on: u16) -> Self {
        IterationRecord { iterations_on }
    }
}

impl Default for IterationRecord {
    fn default() -> Self {
        IterationRecord { iterations_on: 0 }
    }
}

impl BiffRecord for IterationRecord {
    fn id(&self) -> u16 {
        0x011
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(2);
        buf.extend_from_slice(&self.iterations_on.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_iteration_record_id() {
        let record = IterationRecord::default();
        assert_eq!(record.id(), 0x011);
    }

    #[test]
    fn test_iteration_record_default() {
        let record = IterationRecord::default();
        assert_eq!(record.iterations_on, 0);
    }
}
