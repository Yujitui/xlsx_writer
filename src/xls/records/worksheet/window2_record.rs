use crate::xls::records::BiffRecord;

/// Window2Record 记录
///
/// 作用：存储工作表的窗口显示选项
///
/// Window2Record是Excel BIFF格式中的窗口2记录（ID: 0x023E），用于定义工作表的
/// 显示选项，如是否显示网格线、是否冻结窗格、缩放比例等。
///
/// ## 参数说明
///
/// - `options`: 选项标志
///   - bit 0 (0x0001): 显示公式
///   - bit 1 (0x0002): 显示网格线
///   - bit 2 (0x0004): 显示零值
///   - bit 3 (0x0008): 冻结窗格
///   - bit 4 (0x0010): 显示标题
///   - 等等
/// - `first_visible_row`: 第一个可见行
/// - `first_visible_col`: 第一个可见列
/// - `grid_colour`: 网格线颜色索引
/// - `preview_magn`: 打印预览缩放比例
/// - `normal_magn`: 正常视图缩放比例（0表示使用默认值）
#[derive(Debug)]
pub struct Window2Record {
    options: u16,
    first_visible_row: u16,
    first_visible_col: u16,
    grid_colour: u16,
    preview_magn: u16,
    normal_magn: u16,
}

impl Window2Record {
    pub fn new(
        options: u16,
        first_visible_row: u16,
        first_visible_col: u16,
        grid_colour: u16,
        preview_magn: u16,
        normal_magn: u16,
    ) -> Self {
        Window2Record {
            options,
            first_visible_row,
            first_visible_col,
            grid_colour,
            preview_magn,
            normal_magn,
        }
    }
}

impl Default for Window2Record {
    fn default() -> Self {
        Window2Record {
            options: 0x02B6,
            first_visible_row: 0,
            first_visible_col: 0,
            grid_colour: 64,
            preview_magn: 0,
            normal_magn: 0,
        }
    }
}

impl BiffRecord for Window2Record {
    fn id(&self) -> u16 {
        0x023E
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(18);
        buf.extend_from_slice(&self.options.to_le_bytes());
        buf.extend_from_slice(&self.first_visible_row.to_le_bytes());
        buf.extend_from_slice(&self.first_visible_col.to_le_bytes());
        buf.extend_from_slice(&self.grid_colour.to_le_bytes());
        buf.extend_from_slice(&0x0000u16.to_le_bytes());
        buf.extend_from_slice(&self.preview_magn.to_le_bytes());
        buf.extend_from_slice(&self.normal_magn.to_le_bytes());
        buf.extend_from_slice(&0x00000000u32.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_window2_record_id() {
        let record = Window2Record::default();
        assert_eq!(record.id(), 0x023E);
    }

    #[test]
    fn test_window2_record_default_options() {
        let record = Window2Record::default();
        assert_eq!(record.options, 0x02B6);
    }

    #[test]
    fn test_window2_record_data_size() {
        let record = Window2Record::default();
        assert_eq!(record.data().len(), 18);
    }

    #[test]
    fn test_window2_record_default_fields() {
        let record = Window2Record::default();
        assert_eq!(record.first_visible_row, 0);
        assert_eq!(record.first_visible_col, 0);
        assert_eq!(record.grid_colour, 64);
        assert_eq!(record.preview_magn, 0);
        assert_eq!(record.normal_magn, 0);
    }
}
