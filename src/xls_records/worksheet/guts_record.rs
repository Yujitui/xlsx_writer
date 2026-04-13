use crate::xls_records::BiffRecord;

/// GutsRecord 记录
///
/// 作用：存储工作表的行/列分组（大纲）信息
///
/// GutsRecord是Excel BIFF格式中的分组记录（ID: 0x0080），用于定义工作表中
/// 行和列的分组（大纲）级别信息。当工作表使用了行/列分组功能时需要此记录。
///
/// ## 参数说明
///
/// - `row_gut_width`: 行分组区域宽度
/// - `col_gut_height`: 列分组区域高度
/// - `row_visible_levels`: 行可见级别数（大纲级别）
/// - `col_visible_levels`: 列可见级别数（大纲级别）
#[derive(Debug)]
pub struct GutsRecord {
    row_gut_width: u16,
    col_gut_height: u16,
    row_visible_levels: u16,
    col_visible_levels: u16,
}

impl GutsRecord {
    pub fn new(
        row_gut_width: u16,
        col_gut_height: u16,
        row_visible_levels: u16,
        col_visible_levels: u16,
    ) -> Self {
        GutsRecord {
            row_gut_width,
            col_gut_height,
            row_visible_levels,
            col_visible_levels,
        }
    }
}

impl Default for GutsRecord {
    fn default() -> Self {
        GutsRecord {
            row_gut_width: 0,
            col_gut_height: 0,
            row_visible_levels: 0,
            col_visible_levels: 0,
        }
    }
}

impl BiffRecord for GutsRecord {
    fn id(&self) -> u16 {
        0x0080
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(8);
        buf.extend_from_slice(&self.row_gut_width.to_le_bytes());
        buf.extend_from_slice(&self.col_gut_height.to_le_bytes());
        buf.extend_from_slice(&self.row_visible_levels.to_le_bytes());
        buf.extend_from_slice(&self.col_visible_levels.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_guts_record_id() {
        let record = GutsRecord::default();
        assert_eq!(record.id(), 0x0080);
    }

    #[test]
    fn test_guts_record_default() {
        let record = GutsRecord::default();
        assert_eq!(record.row_gut_width, 0);
        assert_eq!(record.col_gut_height, 0);
        assert_eq!(record.row_visible_levels, 0);
        assert_eq!(record.col_visible_levels, 0);
    }

    #[test]
    fn test_guts_record_data_size() {
        let record = GutsRecord::default();
        assert_eq!(record.data().len(), 8);
    }
}
