use crate::xls_records::BiffRecord;

/// VerticalPageBreak 结构体
///
/// ## 作用
///
/// 表示一个垂直分页符（列分页符）。
///
/// ## 参数说明
///
/// - `col`: 分页列索引
/// - `row_start`: 起始行索引
/// - `row_end`: 结束行索引
#[derive(Debug, Clone)]
pub struct VerticalPageBreak {
    pub col: u16,
    pub row_start: u16,
    pub row_end: u16,
}

/// VerticalPageBreaksRecord 记录
///
/// 作用：存储工作表的垂直分页符（列分页符）
///
/// VerticalPageBreaksRecord是Excel BIFF格式中的垂直分页符记录（ID: 0x001A），
/// 用于定义在哪些列处进行分页打印。
///
/// ## 参数说明
///
/// - `breaks`: 分页符向量
#[derive(Debug)]
pub struct VerticalPageBreaksRecord {
    breaks: Vec<VerticalPageBreak>,
}

impl VerticalPageBreaksRecord {
    pub fn new(breaks: Vec<VerticalPageBreak>) -> Self {
        VerticalPageBreaksRecord { breaks }
    }
}

impl Default for VerticalPageBreaksRecord {
    fn default() -> Self {
        VerticalPageBreaksRecord { breaks: Vec::new() }
    }
}

impl BiffRecord for VerticalPageBreaksRecord {
    fn id(&self) -> u16 {
        0x001A
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::new();
        buf.extend_from_slice(&(self.breaks.len() as u16).to_le_bytes());
        for page_break in &self.breaks {
            buf.extend_from_slice(&page_break.col.to_le_bytes());
            buf.extend_from_slice(&page_break.row_start.to_le_bytes());
            buf.extend_from_slice(&page_break.row_end.to_le_bytes());
        }
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_vertical_page_breaks_record_id() {
        let record = VerticalPageBreaksRecord::default();
        assert_eq!(record.id(), 0x001A);
    }

    #[test]
    fn test_vertical_page_breaks_record_default() {
        let record = VerticalPageBreaksRecord::default();
        assert!(record.breaks.is_empty());
    }

    #[test]
    fn test_vertical_page_breaks_record_data_size() {
        let record = VerticalPageBreaksRecord::default();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_vertical_page_breaks_record_with_breaks() {
        let breaks = vec![
            VerticalPageBreak {
                col: 5,
                row_start: 0,
                row_end: 65535,
            },
            VerticalPageBreak {
                col: 10,
                row_start: 0,
                row_end: 65535,
            },
        ];
        let record = VerticalPageBreaksRecord::new(breaks);
        assert_eq!(record.data().len(), 14);
    }
}
