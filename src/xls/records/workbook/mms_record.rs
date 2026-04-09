use super::BiffRecord;

#[derive(Debug, Default)]
pub struct MMSRecord;

impl MMSRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for MMSRecord {
    fn id(&self) -> u16 {
        0x00C1 // MMS record ID
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00] // 2 bytes: 0x0000
    }
}
