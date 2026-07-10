use super::BiffRecord;

#[derive(Debug, Default)]
pub struct Unknown01C0Record;

impl Unknown01C0Record {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for Unknown01C0Record {
    fn id(&self) -> u16 {
        0x01C0
    }

    fn data(&self) -> Vec<u8> {
        Vec::new()
    }
}
