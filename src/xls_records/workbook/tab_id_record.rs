use super::BiffRecord;

/// TabIDRecord 记录
///
/// 作用：定义工作表标签ID列表
///
/// TabIDRecord是Excel BIFF格式中的工作表ID记录（ID: 0x013D），用于存储
/// 工作簿中所有工作表的标签ID列表。这些ID用于标识工作表。
///
/// ## 参数说明
///
/// - `sheet_count`: 工作表数量
/// - 数据：每个工作表的ID（从1开始连续编号）
#[derive(Debug)]
pub struct TabIDRecord {
    sheet_count: u16,
}

impl TabIDRecord {
    pub fn new(sheet_count: u16) -> Self {
        Self { sheet_count }
    }
}

impl BiffRecord for TabIDRecord {
    fn id(&self) -> u16 {
        0x013D // TabID record ID
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity((self.sheet_count as usize) * 2);
        for i in 0..self.sheet_count {
            buf.extend_from_slice(&(i + 1).to_le_bytes()); // Sheet IDs start from 1
        }
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tab_id_record_id() {
        let record = TabIDRecord::new(1);
        assert_eq!(record.id(), 0x013D);
    }

    #[test]
    fn test_tab_id_record_data_size() {
        let record = TabIDRecord::new(3);
        assert_eq!(record.data().len(), 6);
    }

    #[test]
    fn test_tab_id_record_single_sheet() {
        let record = TabIDRecord::new(1);
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }

    #[test]
    fn test_tab_id_record_multiple_sheets() {
        let record = TabIDRecord::new(3);
        let data = record.data();
        assert_eq!(data.len(), 6);
        assert_eq!(&data[0..2], &1u16.to_le_bytes());
        assert_eq!(&data[2..4], &2u16.to_le_bytes());
        assert_eq!(&data[4..6], &3u16.to_le_bytes());
    }
}
