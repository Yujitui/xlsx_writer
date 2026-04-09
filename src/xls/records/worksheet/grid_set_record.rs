use crate::xls::records::BiffRecord;

/// GridSetRecord 记录
///
/// 作用：控制网格线设置标志
///
/// GridSetRecord是Excel BIFF格式中的网格设置记录（ID: 0x0082），用于
/// 标记用户是否设置过网格线。
///
/// ## 参数说明
///
/// - `value`: 网格线设置标志
///   - 1 = 已设置网格线（默认）
///   - 0 = 使用默认设置
#[derive(Debug)]
pub struct GridSetRecord {
    value: u16,
}

impl GridSetRecord {
    pub fn new(value: u16) -> Self {
        GridSetRecord { value }
    }
}

impl Default for GridSetRecord {
    fn default() -> Self {
        GridSetRecord { value: 1 }
    }
}

impl BiffRecord for GridSetRecord {
    fn id(&self) -> u16 {
        0x0082
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(2);
        buf.extend_from_slice(&self.value.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_grid_set_record_id() {
        let record = GridSetRecord::default();
        assert_eq!(record.id(), 0x0082);
    }

    #[test]
    fn test_grid_set_record_default() {
        let record = GridSetRecord::default();
        assert_eq!(record.value, 1);
    }

    #[test]
    fn test_grid_set_record_data_size() {
        let record = GridSetRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
