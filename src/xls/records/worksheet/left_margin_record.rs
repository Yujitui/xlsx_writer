use crate::xls::records::BiffRecord;
use std::mem::size_of;

/// LeftMarginRecord 记录
///
/// 作用：存储页面打印的左边距
///
/// LeftMarginRecord是Excel BIFF格式中的左边距记录（ID: 0x0026），用于定义
/// 打印页面时的左边距值（单位：英寸）。
///
/// ## 参数说明
///
/// - `value`: 左边距值（英寸），默认值为0.3（约7.6毫米）
#[derive(Debug)]
pub struct LeftMarginRecord {
    value: f64,
}

impl LeftMarginRecord {
    pub fn new(value: f64) -> Self {
        LeftMarginRecord { value }
    }
}

impl Default for LeftMarginRecord {
    fn default() -> Self {
        LeftMarginRecord { value: 0.3 }
    }
}

impl BiffRecord for LeftMarginRecord {
    fn id(&self) -> u16 {
        0x0026
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
    fn test_left_margin_record_id() {
        let record = LeftMarginRecord::default();
        assert_eq!(record.id(), 0x0026);
    }

    #[test]
    fn test_left_margin_record_default() {
        let record = LeftMarginRecord::default();
        assert_eq!(record.value, 0.3);
    }

    #[test]
    fn test_left_margin_record_data_size() {
        let record = LeftMarginRecord::default();
        assert_eq!(record.data().len(), 8);
    }
}
