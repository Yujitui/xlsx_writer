use crate::xls_records::BiffRecord;

#[derive(Debug)]
pub struct DefColWidthRecord {
    width: u16,
}

impl Default for DefColWidthRecord {
    fn default() -> Self {
        DefColWidthRecord { width: 8 }
    }
}

impl BiffRecord for DefColWidthRecord {
    fn id(&self) -> u16 {
        0x0055
    }

    fn data(&self) -> Vec<u8> {
        self.width.to_le_bytes().to_vec()
    }
}
