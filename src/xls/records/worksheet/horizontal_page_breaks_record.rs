use crate::xls::records::BiffRecord;

/// PageBreak 结构体
///
/// ## 作用
///
/// 表示一个水平分页符（行分页符）。
///
/// ## 参数说明
///
/// - `row`: 分页行索引
/// - `col_start`: 起始列索引
/// - `col_end`: 结束列索引
#[derive(Debug, Clone)]
pub struct PageBreak {
    pub row: u16,
    pub col_start: u16,
    pub col_end: u16,
}

/// HorizontalPageBreaksRecord 记录
///
/// 作用：存储工作表的水平分页符（行分页符）
///
/// HorizontalPageBreaksRecord是Excel BIFF格式中的水平分页符记录（ID: 0x001B），
/// 用于定义在哪些行处进行分页打印。
///
/// ## 参数说明
///
/// - `breaks`: 分页符向量
#[derive(Debug)]
pub struct HorizontalPageBreaksRecord {
    breaks: Vec<PageBreak>,
}

impl HorizontalPageBreaksRecord {
    pub fn new(breaks: Vec<PageBreak>) -> Self {
        HorizontalPageBreaksRecord { breaks }
    }
}

impl Default for HorizontalPageBreaksRecord {
    fn default() -> Self {
        HorizontalPageBreaksRecord { breaks: Vec::new() }
    }
}

impl BiffRecord for HorizontalPageBreaksRecord {
    fn id(&self) -> u16 {
        0x001B
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::new();
        buf.extend_from_slice(&(self.breaks.len() as u16).to_le_bytes());
        for page_break in &self.breaks {
            buf.extend_from_slice(&page_break.row.to_le_bytes());
            buf.extend_from_slice(&page_break.col_start.to_le_bytes());
            buf.extend_from_slice(&page_break.col_end.to_le_bytes());
        }
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_horizontal_page_breaks_record_id() {
        let record = HorizontalPageBreaksRecord::default();
        assert_eq!(record.id(), 0x001B);
    }

    #[test]
    fn test_horizontal_page_breaks_record_default() {
        let record = HorizontalPageBreaksRecord::default();
        assert!(record.breaks.is_empty());
    }

    #[test]
    fn test_horizontal_page_breaks_record_data_size() {
        let record = HorizontalPageBreaksRecord::default();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_horizontal_page_breaks_record_with_breaks() {
        let breaks = vec![
            PageBreak {
                row: 10,
                col_start: 0,
                col_end: 255,
            },
            PageBreak {
                row: 20,
                col_start: 0,
                col_end: 255,
            },
        ];
        let record = HorizontalPageBreaksRecord::new(breaks);
        assert_eq!(record.data().len(), 14);
    }
}
