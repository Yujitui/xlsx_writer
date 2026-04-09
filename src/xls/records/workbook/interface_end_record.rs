use super::BiffRecord;

#[derive(Debug, Default)]
pub struct InterfaceEndRecord;

impl InterfaceEndRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for InterfaceEndRecord {
    fn id(&self) -> u16 {
        0x00E2 // Interface End record ID
    }

    fn data(&self) -> Vec<u8> {
        Vec::new() // Empty data
    }
}
