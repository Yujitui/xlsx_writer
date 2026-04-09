use super::BiffRecord;

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
