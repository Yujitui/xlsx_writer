use super::BiffRecord;

/// Window1Record 记录
///
/// 作用：定义工作簿窗口的布局
///
/// Window1Record是Excel BIFF格式中的窗口记录（ID: 0x003D），用于定义
/// 工作簿窗口的位置、大小和显示选项。
///
/// ## 参数说明
///
/// - `hpos_twips`: 窗口水平位置（单位：twips，1/1440英寸）
/// - `vpos_twips`: 窗口垂直位置（单位：twips）
/// - `width_twips`: 窗口宽度（单位：twips）
/// - `height_twips`: 窗口高度（单位：twips）
/// - `flags`: 窗口选项标志
/// - `active_sheet`: 当前活动工作表索引
/// - `first_tab_index`: 第一个显示的工作表标签索引
/// - `selected_tabs`: 选中的工作表标签数量
/// - `tab_width`: 工作表标签栏宽度
#[derive(Debug)]
pub struct Window1Record {
    hpos_twips: u16,
    vpos_twips: u16,
    width_twips: u16,
    height_twips: u16,
    flags: u16,
    active_sheet: u16,
    first_tab_index: u16,
    selected_tabs: u16,
    tab_width: u16,
}

impl Window1Record {
    pub fn new(
        hpos_twips: u16,
        vpos_twips: u16,
        width_twips: u16,
        height_twips: u16,
        flags: u16,
        active_sheet: u16,
        first_tab_index: u16,
        selected_tabs: u16,
        tab_width: u16,
    ) -> Self {
        Window1Record {
            hpos_twips,
            vpos_twips,
            width_twips,
            height_twips,
            flags,
            active_sheet,
            first_tab_index,
            selected_tabs,
            tab_width,
        }
    }
}

impl Default for Window1Record {
    fn default() -> Self {
        Window1Record {
            hpos_twips: 0x01E0,
            vpos_twips: 0x005A,
            width_twips: 0x3FCF,
            height_twips: 0x2A4E,
            flags: 0x0038,
            active_sheet: 0,
            first_tab_index: 0,
            selected_tabs: 1,
            tab_width: 0x0258,
        }
    }
}

impl BiffRecord for Window1Record {
    fn id(&self) -> u16 {
        0x003D // Window1 record ID
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(18);
        buf.extend_from_slice(&self.hpos_twips.to_le_bytes());
        buf.extend_from_slice(&self.vpos_twips.to_le_bytes());
        buf.extend_from_slice(&self.width_twips.to_le_bytes());
        buf.extend_from_slice(&self.height_twips.to_le_bytes());
        buf.extend_from_slice(&self.flags.to_le_bytes());
        buf.extend_from_slice(&self.active_sheet.to_le_bytes());
        buf.extend_from_slice(&self.first_tab_index.to_le_bytes());
        buf.extend_from_slice(&self.selected_tabs.to_le_bytes());
        buf.extend_from_slice(&self.tab_width.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xls_records::BiffRecord;

    #[test]
    fn test_window1_record_id() {
        let record = Window1Record::new(0, 0, 0, 0, 0, 0, 0, 0, 0);
        assert_eq!(record.id(), 0x003D);
    }

    #[test]
    fn test_window1_record_data_size() {
        let record = Window1Record::new(0, 0, 0, 0, 0, 0, 0, 0, 0);
        assert_eq!(record.data().len(), 18); // 9 * 2 bytes
    }

    #[test]
    fn test_window1_record_serialize() {
        let record = Window1Record::new(0x01E0, 0x005A, 0x3FCF, 0x2A4E, 0x0038, 0, 0, 1, 0x0258);
        let serialized = record.serialize();

        // Check record header: ID (2 bytes) + length (2 bytes)
        assert_eq!(serialized[0], 0x3D);
        assert_eq!(serialized[1], 0x00);
        assert_eq!(serialized[2], 0x12);
        assert_eq!(serialized[3], 0x00);

        // Check data
        assert_eq!(serialized[4..6], 0x01E0u16.to_le_bytes());
        assert_eq!(serialized[6..8], 0x005Au16.to_le_bytes());
    }

    #[test]
    fn test_window1_record_default() {
        let record = Window1Record::default();

        // Verify default values match xlwt
        assert_eq!(record.hpos_twips, 0x01E0);
        assert_eq!(record.vpos_twips, 0x005A);
        assert_eq!(record.width_twips, 0x3FCF);
        assert_eq!(record.height_twips, 0x2A4E);
        assert_eq!(record.flags, 0x0038);
        assert_eq!(record.active_sheet, 0);
        assert_eq!(record.first_tab_index, 0);
        assert_eq!(record.selected_tabs, 1);
        assert_eq!(record.tab_width, 0x0258);
    }
}
