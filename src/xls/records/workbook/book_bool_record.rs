use super::BiffRecord;

#[derive(Debug)]
pub struct BookBoolRecord;

impl BookBoolRecord {
    pub fn new() -> Self {
        BookBoolRecord
    }
}

impl BiffRecord for BookBoolRecord {
    fn id(&self) -> u16 {
        0x00DA
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_book_bool_record_id() {
        let record = BookBoolRecord::new();
        assert_eq!(record.id(), 0x00DA);
    }

    #[test]
    fn test_book_bool_record_data_size() {
        let record = BookBoolRecord::new();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_book_bool_record_data() {
        let record = BookBoolRecord::new();
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }
}
