use crate::xls::records::BiffRecord;
use std::mem::size_of;

/// TopMarginRecord 记录
///
/// 作用：存储页面打印的上边距
///
/// TopMarginRecord是Excel BIFF格式中的上边距记录（ID: 0x0028），用于定义
/// 打印页面时的上边距值（单位：英寸）。
///
/// ## 参数说明
///
/// - `value`: 上边距值（英寸），默认值为0.61（约15.5毫米）
#[derive(Debug)]
pub struct TopMarginRecord {
    value: f64,
}

impl TopMarginRecord {
    pub fn new(value: f64) -> Self {
        TopMarginRecord { value }
    }
}

impl Default for TopMarginRecord {
    fn default() -> Self {
        TopMarginRecord { value: 0.61 }
    }
}

impl BiffRecord for TopMarginRecord {
    fn id(&self) -> u16 {
        0x0028
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
    fn test_top_margin_record_id() {
        let record = TopMarginRecord::default();
        assert_eq!(record.id(), 0x0028);
    }

    #[test]
    fn test_top_margin_record_default() {
        let record = TopMarginRecord::default();
        assert_eq!(record.value, 0.61);
    }

    #[test]
    fn test_top_margin_record_data_size() {
        let record = TopMarginRecord::default();
        assert_eq!(record.data().len(), 8);
    }
}
