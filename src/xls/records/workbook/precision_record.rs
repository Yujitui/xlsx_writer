use super::BiffRecord;

#[derive(Debug)]
pub struct PrecisionRecord {
    use_real_values: bool,
}

impl PrecisionRecord {
    pub fn new(use_real_values: bool) -> Self {
        PrecisionRecord { use_real_values }
    }
}

impl Default for PrecisionRecord {
    fn default() -> Self {
        PrecisionRecord::new(true)
    }
}

impl BiffRecord for PrecisionRecord {
    fn id(&self) -> u16 {
        0x000E
    }

    fn data(&self) -> Vec<u8> {
        let value: u16 = if self.use_real_values { 1 } else { 0 };
        value.to_le_bytes().to_vec()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_precision_record_id() {
        let record = PrecisionRecord::new(false);
        assert_eq!(record.id(), 0x000E);
    }

    #[test]
    fn test_precision_record_data_size() {
        let record = PrecisionRecord::new(false);
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_precision_use_real_values_false() {
        let record = PrecisionRecord::new(false);
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }

    #[test]
    fn test_precision_use_real_values_true() {
        let record = PrecisionRecord::new(true);
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }
}
