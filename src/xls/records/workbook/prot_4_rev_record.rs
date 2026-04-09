use super::BiffRecord;

#[derive(Debug, Clone, PartialEq)]
pub struct Prot4RevRecord;

impl Prot4RevRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for Prot4RevRecord {
    fn id(&self) -> u16 {
        0x01AF
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00]
    }
}
