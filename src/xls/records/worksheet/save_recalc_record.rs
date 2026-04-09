use crate::xls::records::BiffRecord;

/// SaveRecalcRecord 记录
///
/// 作用：控制工作表的重算标志
///
/// SaveRecalcRecord是Excel BIFF格式中的保存重算记录（ID: 0x05F），用于指示
/// Excel在打开文件时是否需要重新计算公式。
///
/// ## 参数说明
///
/// - `recalc`: 重算标志
///   - 0 = 打开时不需要重新计算
///   - 1 = 打开时需要重新计算
#[derive(Debug)]
pub struct SaveRecalcRecord {
    recalc: u16,
}

impl SaveRecalcRecord {
    pub fn new(recalc: u16) -> Self {
        SaveRecalcRecord { recalc }
    }
}

impl Default for SaveRecalcRecord {
    fn default() -> Self {
        SaveRecalcRecord { recalc: 0 }
    }
}

impl BiffRecord for SaveRecalcRecord {
    fn id(&self) -> u16 {
        0x05F
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(2);
        buf.extend_from_slice(&self.recalc.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_save_recalc_record_id() {
        let record = SaveRecalcRecord::default();
        assert_eq!(record.id(), 0x05F);
    }

    #[test]
    fn test_save_recalc_record_default() {
        let record = SaveRecalcRecord::default();
        assert_eq!(record.recalc, 0);
    }
}
