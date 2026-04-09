use super::BiffRecord;

#[derive(Debug)]
pub struct UseSelfsRecord;

impl UseSelfsRecord {
    pub fn new() -> Self {
        UseSelfsRecord
    }
}

impl BiffRecord for UseSelfsRecord {
    fn id(&self) -> u16 {
        0x0160
    }

    fn data(&self) -> Vec<u8> {
        vec![0x01, 0x00]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_use_selfs_record_id() {
        let record = UseSelfsRecord::new();
        assert_eq!(record.id(), 0x0160);
    }

    #[test]
    fn test_use_selfs_record_data_size() {
        let record = UseSelfsRecord::new();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_use_selfs_record_data() {
        let record = UseSelfsRecord::new();
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }
}
