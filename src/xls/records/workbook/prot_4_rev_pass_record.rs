use super::BiffRecord;

#[derive(Debug, Clone, PartialEq)]
pub struct Prot4RevPassRecord;

impl Prot4RevPassRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for Prot4RevPassRecord {
    fn id(&self) -> u16 {
        0x01BC
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00]
    }
}
