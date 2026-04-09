use super::BiffRecord;

#[derive(Debug, Default)]
pub struct DSFRecord;

impl DSFRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for DSFRecord {
    fn id(&self) -> u16 {
        0x0161 // DSF record ID
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00] // Hard-coded 0x0000
    }
}
