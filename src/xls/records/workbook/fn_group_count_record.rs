use super::BiffRecord;

#[derive(Debug, Default)]
pub struct FnGroupCountRecord;

impl FnGroupCountRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for FnGroupCountRecord {
    fn id(&self) -> u16 {
        0x009C // FnGroupCount record ID
    }

    fn data(&self) -> Vec<u8> {
        vec![0x0E, 0x00] // Hard-coded 0x000E (14 function groups)
    }
}
