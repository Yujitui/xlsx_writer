// src/biff_records/codepage_record.rs

use super::BiffRecord;

#[derive(Debug, Default)]
pub struct CodepageRecord;

impl CodepageRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for CodepageRecord {
    fn id(&self) -> u16 {
        0x0042 // Codepage record ID
    }

    fn data(&self) -> Vec<u8> {
        vec![0xB0, 0x04] // Hard-coded UTF-16 (1200 in little-endian)
    }
}
