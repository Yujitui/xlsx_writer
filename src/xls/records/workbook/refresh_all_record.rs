use super::BiffRecord;

#[derive(Debug)]
pub struct RefreshAllRecord;

impl RefreshAllRecord {
    pub fn new() -> Self {
        RefreshAllRecord
    }
}

impl BiffRecord for RefreshAllRecord {
    fn id(&self) -> u16 {
        0x01B7
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_refresh_all_record_id() {
        let record = RefreshAllRecord::new();
        assert_eq!(record.id(), 0x01B7);
    }

    #[test]
    fn test_refresh_all_record_data_size() {
        let record = RefreshAllRecord::new();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_refresh_all_record_data() {
        let record = RefreshAllRecord::new();
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }
}
