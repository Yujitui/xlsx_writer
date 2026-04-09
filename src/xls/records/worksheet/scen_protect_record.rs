use crate::xls::records::BiffRecord;

/// ScenProtectRecord 记录
///
/// 作用：控制场景保护状态
///
/// ScenProtectRecord是Excel BIFF格式中的场景保护记录（ID: 0x00DD），用于
/// 保护工作表中的方案（Scenario）。方案允许用户定义不同输入值组合并查看结果。
///
/// ## 参数说明
///
/// - `value`: 保护标志
///   - 0 = 未保护（默认）
///   - 1 = 已保护
#[derive(Debug)]
pub struct ScenProtectRecord {
    value: u16,
}

impl ScenProtectRecord {
    pub fn new(value: u16) -> Self {
        ScenProtectRecord { value }
    }
}

impl Default for ScenProtectRecord {
    fn default() -> Self {
        ScenProtectRecord { value: 0 }
    }
}

impl BiffRecord for ScenProtectRecord {
    fn id(&self) -> u16 {
        0x00DD
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
    fn test_scen_protect_record_id() {
        let record = ScenProtectRecord::default();
        assert_eq!(record.id(), 0x00DD);
    }

    #[test]
    fn test_scen_protect_record_default() {
        let record = ScenProtectRecord::default();
        assert_eq!(record.value, 0);
    }

    #[test]
    fn test_scen_protect_record_data_size() {
        let record = ScenProtectRecord::default();
        assert_eq!(record.data().len(), 2);
    }
}
