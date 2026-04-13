use crate::xls_records::BiffRecord;
use std::mem::size_of;

/// RightMarginRecord 记录
///
/// 作用：存储页面打印的右边距
///
/// RightMarginRecord是Excel BIFF格式中的右边距记录（ID: 0x0027），用于定义
/// 打印页面时的右边距值（单位：英寸）。
///
/// ## 参数说明
///
/// - `value`: 右边距值（英寸），默认值为0.3（约7.6毫米）
#[derive(Debug)]
pub struct RightMarginRecord {
    value: f64,
}

impl RightMarginRecord {
    pub fn new(value: f64) -> Self {
        RightMarginRecord { value }
    }
}

impl Default for RightMarginRecord {
    fn default() -> Self {
        RightMarginRecord { value: 0.3 }
    }
}

impl BiffRecord for RightMarginRecord {
    fn id(&self) -> u16 {
        0x0027
    }

    //noinspection DuplicatedCode
    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(size_of::<f64>());
        buf.extend_from_slice(&self.value.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_right_margin_record_id() {
        let record = RightMarginRecord::default();
        assert_eq!(record.id(), 0x0027);
    }

    #[test]
    fn test_right_margin_record_default() {
        let record = RightMarginRecord::default();
        assert_eq!(record.value, 0.3);
    }

    #[test]
    fn test_right_margin_record_data_size() {
        let record = RightMarginRecord::default();
        assert_eq!(record.data().len(), 8);
    }
}
