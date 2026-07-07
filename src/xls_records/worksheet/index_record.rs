use crate::xls_records::BiffRecord;

#[derive(Debug)]
pub struct IndexRecord {
    rw_mic: u32,
    rw_mac: u32,
    ib_xf: u32,
    rgib_rw: Vec<u32>,
}

impl IndexRecord {
    pub fn new(rw_mic: u32, rw_mac: u32, ib_xf: u32, rgib_rw: Vec<u32>) -> Self {
        IndexRecord {
            rw_mic,
            rw_mac,
            ib_xf,
            rgib_rw,
        }
    }
}

impl BiffRecord for IndexRecord {
    fn id(&self) -> u16 {
        0x020B
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(16 + self.rgib_rw.len() * 4);
        buf.extend_from_slice(&[0u8; 4]); // reserved
        buf.extend_from_slice(&self.rw_mic.to_le_bytes());
        buf.extend_from_slice(&self.rw_mac.to_le_bytes());
        buf.extend_from_slice(&self.ib_xf.to_le_bytes());
        for fp in &self.rgib_rw {
            buf.extend_from_slice(&fp.to_le_bytes());
        }
        buf
    }
}
