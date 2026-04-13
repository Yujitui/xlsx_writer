use crate::xls_records::BiffRecord;

/// SetupPageRecord 记录
///
/// 作用：存储页面打印设置
///
/// SetupPageRecord是Excel BIFF格式中的页面设置记录（ID: 0x00A1），用于定义
/// 工作表的完整打印页面设置，包括纸张大小、缩放、页边距等。
///
/// ## 参数说明
///
/// - `paper_size`: 纸张大小代码（如9=A4, 1=Letter）
/// - `print_scaling`: 打印缩放百分比（100=100%）
/// - `start_page_number`: 起始页码（1=自动）
/// - `fit_width_to_pages`: 缩放到指定页宽（0=不使用）
/// - `fit_height_to_pages`: 缩放到指定页高（0=不使用）
/// - `options`: 选项标志
/// - `print_hres`: 水平打印分辨率（dpi）
/// - `print_vres`: 垂直打印分辨率（dpi）
/// - `header_margin`: 页眉边距（英寸）
/// - `footer_margin`: 页脚边距（英寸）
/// - `copies_num`: 打印份数
#[derive(Debug)]
pub struct SetupPageRecord {
    paper_size: u16,
    print_scaling: u16,
    start_page_number: u16,
    fit_width_to_pages: u16,
    fit_height_to_pages: u16,
    options: u16,
    print_hres: u32,
    print_vres: u32,
    header_margin: f64,
    footer_margin: f64,
    copies_num: u16,
}

impl SetupPageRecord {
    pub fn new(
        paper_size: u16,
        print_scaling: u16,
        start_page_number: u16,
        fit_width_to_pages: u16,
        fit_height_to_pages: u16,
        options: u16,
        print_hres: u32,
        print_vres: u32,
        header_margin: f64,
        footer_margin: f64,
        copies_num: u16,
    ) -> Self {
        SetupPageRecord {
            paper_size,
            print_scaling,
            start_page_number,
            fit_width_to_pages,
            fit_height_to_pages,
            options,
            print_hres,
            print_vres,
            header_margin,
            footer_margin,
            copies_num,
        }
    }
}

impl Default for SetupPageRecord {
    fn default() -> Self {
        SetupPageRecord {
            paper_size: 9,
            print_scaling: 100,
            start_page_number: 1,
            fit_width_to_pages: 1,
            fit_height_to_pages: 1,
            options: 0x0083,
            print_hres: 0x012C,
            print_vres: 0x012C,
            header_margin: 0.1,
            footer_margin: 0.1,
            copies_num: 1,
        }
    }
}

impl BiffRecord for SetupPageRecord {
    fn id(&self) -> u16 {
        0x00A1
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(34);
        buf.extend_from_slice(&self.paper_size.to_le_bytes());
        buf.extend_from_slice(&self.print_scaling.to_le_bytes());
        buf.extend_from_slice(&self.start_page_number.to_le_bytes());
        buf.extend_from_slice(&self.fit_width_to_pages.to_le_bytes());
        buf.extend_from_slice(&self.fit_height_to_pages.to_le_bytes());
        buf.extend_from_slice(&self.options.to_le_bytes());
        buf.extend_from_slice(&self.print_hres.to_le_bytes());
        buf.extend_from_slice(&self.print_vres.to_le_bytes());
        buf.extend_from_slice(&self.header_margin.to_le_bytes());
        buf.extend_from_slice(&self.footer_margin.to_le_bytes());
        buf.extend_from_slice(&self.copies_num.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_setup_page_record_id() {
        let record = SetupPageRecord::default();
        assert_eq!(record.id(), 0x00A1);
    }

    #[test]
    fn test_setup_page_record_default() {
        let record = SetupPageRecord::default();
        assert_eq!(record.paper_size, 9);
        assert_eq!(record.print_scaling, 100);
        assert_eq!(record.start_page_number, 1);
        assert_eq!(record.fit_width_to_pages, 1);
        assert_eq!(record.fit_height_to_pages, 1);
        assert_eq!(record.options, 0x0083);
        assert_eq!(record.print_hres, 0x012C);
        assert_eq!(record.print_vres, 0x012C);
        assert_eq!(record.header_margin, 0.1);
        assert_eq!(record.footer_margin, 0.1);
        assert_eq!(record.copies_num, 1);
    }

    #[test]
    fn test_setup_page_record_data_size() {
        let record = SetupPageRecord::default();
        assert_eq!(record.data().len(), 38);
    }
}
