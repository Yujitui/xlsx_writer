use super::BiffRecord;

#[derive(Debug, Clone)]
pub struct HideObjRecord;

impl HideObjRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for HideObjRecord {
    fn id(&self) -> u16 {
        0x008D
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00]
    }
}
