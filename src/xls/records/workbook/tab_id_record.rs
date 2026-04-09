use super::BiffRecord;

#[derive(Debug)]
pub struct TabIDRecord {
    sheet_count: u16,
}

impl TabIDRecord {
    pub fn new(sheet_count: u16) -> Self {
        Self { sheet_count }
    }
}

impl BiffRecord for TabIDRecord {
    fn id(&self) -> u16 {
        0x013D // TabID record ID
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity((self.sheet_count as usize) * 2);
        for i in 0..self.sheet_count {
            buf.extend_from_slice(&(i + 1).to_le_bytes()); // Sheet IDs start from 1
        }
        buf
    }
}
