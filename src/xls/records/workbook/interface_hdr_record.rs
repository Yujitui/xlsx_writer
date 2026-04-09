use super::BiffRecord;

#[derive(Debug)]
pub struct InterfaceHdrRecord;

impl InterfaceHdrRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for InterfaceHdrRecord {
    fn id(&self) -> u16 {
        0x00E1 // Interface Header record ID
    }

    fn data(&self) -> Vec<u8> {
        vec![0xB0, 0x04] // Corrected: 0x04B0 (1200 decimal)
    }
}
