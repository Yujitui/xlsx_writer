use crate::xls_records::BiffRecord;

/// RefModeRecord 记录
///
/// 作用：设置工作表的引用模式
///
/// RefModeRecord是Excel BIFF格式中的引用模式记录（ID: 0x000F），用于定义
/// 单元格引用在公式中的表示方式（A1样式或R1C1样式）。
///
/// ## 参数说明
///
/// - `ref_mode`: 引用模式
///   - 1 = A1样式（默认，如A1, B2等）
///   - 2 = R1C1样式（如R1C1, R2C3等）
#[derive(Debug)]
pub struct RefModeRecord {
    ref_mode: u16,
}

impl RefModeRecord {
    pub fn new(ref_mode: u16) -> Self {
        RefModeRecord { ref_mode }
    }
}

impl Default for RefModeRecord {
    fn default() -> Self {
        RefModeRecord { ref_mode: 1 }
    }
}

impl BiffRecord for RefModeRecord {
    fn id(&self) -> u16 {
        0x000F
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(2);
        buf.extend_from_slice(&self.ref_mode.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_ref_mode_record_id() {
        let record = RefModeRecord::default();
        assert_eq!(record.id(), 0x000F);
    }

    #[test]
    fn test_ref_mode_record_default() {
        let record = RefModeRecord::default();
        assert_eq!(record.ref_mode, 1);
    }
}
