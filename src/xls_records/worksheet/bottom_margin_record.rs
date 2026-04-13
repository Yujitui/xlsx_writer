use crate::xls_records::BiffRecord;
use std::mem::size_of;

/// BottomMarginRecord 记录
///
/// 作用：存储页面打印的下边距
///
/// BottomMarginRecord是Excel BIFF格式中的下边距记录（ID: 0x0029），用于定义
/// 打印页面时的下边距值（单位：英寸）。
///
/// ## 参数说明
///
/// - `value`: 下边距值（英寸），默认值为0.37（约9.5毫米）
#[derive(Debug)]
pub struct BottomMarginRecord {
    value: f64,
}

impl BottomMarginRecord {
    pub fn new(value: f64) -> Self {
        BottomMarginRecord { value }
    }
}

impl Default for BottomMarginRecord {
    fn default() -> Self {
        BottomMarginRecord { value: 0.37 }
    }
}

impl BiffRecord for BottomMarginRecord {
    fn id(&self) -> u16 {
        0x0029
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
    fn test_bottom_margin_record_id() {
        let record = BottomMarginRecord::default();
        assert_eq!(record.id(), 0x0029);
    }

    #[test]
    fn test_bottom_margin_record_default() {
        let record = BottomMarginRecord::default();
        assert_eq!(record.value, 0.37);
    }

    #[test]
    fn test_bottom_margin_record_data_size() {
        let record = BottomMarginRecord::default();
        assert_eq!(record.data().len(), 8);
    }
}
