use super::BiffRecord;

/// CountryRecord 记录
///
/// 作用：定义工作簿的国家/地区设置
///
/// CountryRecord是Excel BIFF格式中的国家记录（ID: 0x008C），用于定义
/// 工作簿使用的语言和国家/地区设置，用于日期和货币格式等。
///
/// ## 参数说明
///
/// - `ui_id`: 用户界面语言ID
///   - 1 = 英语（美国）
///   - 31 = 简体中文
///   - 1028 = 繁体中文
/// - `sys_settings_id`: 系统设置语言ID
///   - 默认值为 1（美国）
#[derive(Debug)]
pub struct CountryRecord {
    ui_id: u16,
    sys_settings_id: u16,
}

impl CountryRecord {
    pub fn new(ui_id: u16, sys_settings_id: u16) -> Self {
        CountryRecord {
            ui_id,
            sys_settings_id,
        }
    }
}

impl Default for CountryRecord {
    fn default() -> Self {
        CountryRecord::new(1, 1) // Default to USA
    }
}

impl BiffRecord for CountryRecord {
    fn id(&self) -> u16 {
        0x008C
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(4);
        buf.extend_from_slice(&self.ui_id.to_le_bytes());
        buf.extend_from_slice(&self.sys_settings_id.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_country_record_id() {
        let record = CountryRecord::new(1, 1);
        assert_eq!(record.id(), 0x008C);
    }

    #[test]
    fn test_country_record_data_size() {
        let record = CountryRecord::new(1, 1);
        assert_eq!(record.data().len(), 4);
    }

    #[test]
    fn test_country_record_data() {
        let record = CountryRecord::new(1, 1);
        assert_eq!(record.data(), vec![0x01, 0x00, 0x01, 0x00]);
    }
}
