use crate::xls::records::BiffRecord;

/// DefaultRowHeightRecord 记录
///
/// 作用：存储工作表的默认行高设置
///
/// DefaultRowHeightRecord是Excel BIFF格式中的默认行高记录（ID: 0x0225），用于
/// 定义工作表中未明确设置高度的行的默认高度。
///
/// ## 参数说明
///
/// - `options`: 选项标志
///   - bit 0: 是否使用自定义行高（0=使用Excel默认, 1=使用自定义）
/// - `def_height`: 默认行高（单位：1/20点）
///   - 默认值 0x00FF = 12.75 点
#[derive(Debug)]
pub struct DefaultRowHeightRecord {
    options: u16,
    def_height: u16,
}

impl DefaultRowHeightRecord {
    pub fn new(options: u16, def_height: u16) -> Self {
        DefaultRowHeightRecord {
            options,
            def_height,
        }
    }
}

impl Default for DefaultRowHeightRecord {
    fn default() -> Self {
        DefaultRowHeightRecord {
            options: 0x0000,
            def_height: 0x00FF,
        }
    }
}

impl BiffRecord for DefaultRowHeightRecord {
    fn id(&self) -> u16 {
        0x0225
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(4);
        buf.extend_from_slice(&self.options.to_le_bytes());
        buf.extend_from_slice(&self.def_height.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_row_height_record_id() {
        let record = DefaultRowHeightRecord::default();
        assert_eq!(record.id(), 0x0225);
    }

    #[test]
    fn test_default_row_height_record_default() {
        let record = DefaultRowHeightRecord::default();
        assert_eq!(record.options, 0x0000);
        assert_eq!(record.def_height, 0x00FF);
    }

    #[test]
    fn test_default_row_height_record_data_size() {
        let record = DefaultRowHeightRecord::default();
        assert_eq!(record.data().len(), 4);
    }
}
