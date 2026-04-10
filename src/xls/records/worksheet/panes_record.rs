use crate::xls::records::BiffRecord;

/// PanesRecord 记录
///
/// 作用：存储工作表的窗格（拆分窗口）信息
///
/// PanesRecord是Excel BIFF格式中的窗格记录（ID: 0x0041），用于定义工作表的
/// 窗格拆分方式。当工作表使用了"冻结窗格"或"拆分"功能时需要此记录。
///
/// **注意**：当没有窗格拆分时，此记录的serialize方法返回空数据，不写入文件。
///
/// ## 参数说明
///
/// - `vert_split_pos`: 垂直拆分位置（列索引，0表示无垂直拆分）
/// - `horz_split_pos`: 水平拆分位置（行索引，0表示无水平拆分）
/// - `first_visible_row`: 第一个可见行（拆分后）
/// - `first_visible_col`: 第一个可见列（拆分后）
/// - `active_pane`: 活动窗格
///   - 0 = 左上, 1 = 右上, 2 = 左下, 3 = 右下
#[derive(Debug)]
pub struct PanesRecord {
    vert_split_pos: u16,
    horz_split_pos: u16,
    first_visible_row: u16,
    first_visible_col: u16,
    active_pane: u8,
}

impl PanesRecord {
    pub fn new(
        vert_split_pos: u16,
        horz_split_pos: u16,
        first_visible_row: u16,
        first_visible_col: u16,
        active_pane: u8,
    ) -> Self {
        PanesRecord {
            vert_split_pos,
            horz_split_pos,
            first_visible_row,
            first_visible_col,
            active_pane,
        }
    }
}

impl Default for PanesRecord {
    fn default() -> Self {
        PanesRecord {
            vert_split_pos: 0,
            horz_split_pos: 0,
            first_visible_row: 0,
            first_visible_col: 0,
            active_pane: 0,
        }
    }
}

impl BiffRecord for PanesRecord {
    fn id(&self) -> u16 {
        0x0041
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(10);
        buf.extend_from_slice(&self.vert_split_pos.to_le_bytes());
        buf.extend_from_slice(&self.horz_split_pos.to_le_bytes());
        buf.extend_from_slice(&self.first_visible_row.to_le_bytes());
        buf.extend_from_slice(&self.first_visible_col.to_le_bytes());
        buf.push(self.active_pane);
        buf.push(0x00);
        buf
    }

    fn serialize(&self) -> Vec<u8> {
        if self.vert_split_pos == 0 && self.horz_split_pos == 0 {
            return Vec::new();
        }
        let payload = self.data();
        let len = payload.len() as u16;
        let mut result = Vec::with_capacity(4 + payload.len());
        result.extend_from_slice(&self.id().to_le_bytes());
        result.extend_from_slice(&len.to_le_bytes());
        result.extend_from_slice(&payload);
        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_panes_record_id() {
        let record = PanesRecord::default();
        assert_eq!(record.id(), 0x0041);
    }

    #[test]
    fn test_panes_record_default() {
        let record = PanesRecord::default();
        assert_eq!(record.vert_split_pos, 0);
        assert_eq!(record.horz_split_pos, 0);
    }

    #[test]
    fn test_panes_record_data_size() {
        let record = PanesRecord::default();
        assert_eq!(record.data().len(), 10);
    }

    #[test]
    fn test_panes_record_empty_serialize() {
        let record = PanesRecord::default();
        let serialized = record.serialize();
        assert!(serialized.is_empty());
    }
}
