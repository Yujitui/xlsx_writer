use super::BiffRecord;

#[derive(Debug)]
pub struct DateModeRecord {
    from_1904: bool,
}

impl DateModeRecord {
    pub fn new(from_1904: bool) -> Self {
        DateModeRecord { from_1904 }
    }
}

impl Default for DateModeRecord {
    fn default() -> Self {
        Self::new(false)
    }
}

impl BiffRecord for DateModeRecord {
    fn id(&self) -> u16 {
        0x0022
    }

    fn data(&self) -> Vec<u8> {
        let value: u16 = if self.from_1904 { 1 } else { 0 };
        value.to_le_bytes().to_vec()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_date_mode_record_id() {
        let record = DateModeRecord::new(false);
        assert_eq!(record.id(), 0x0022);
    }

    #[test]
    fn test_date_mode_record_data_size() {
        let record = DateModeRecord::new(false);
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_date_mode_from_1904_false() {
        let record = DateModeRecord::new(false);
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }

    #[test]
    fn test_date_mode_from_1904_true() {
        let record = DateModeRecord::new(true);
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }
}
