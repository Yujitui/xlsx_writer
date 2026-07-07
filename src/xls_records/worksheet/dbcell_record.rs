use crate::xls_records::BiffRecord;

#[derive(Debug)]
pub struct DBCellRecord {
    db_rtrw: u32,
    rgdb: Vec<u16>,
}

impl DBCellRecord {
    pub fn new(db_rtrw: u32, rgdb: Vec<u16>) -> Self {
        DBCellRecord { db_rtrw, rgdb }
    }
}

impl BiffRecord for DBCellRecord {
    fn id(&self) -> u16 {
        0x00D7
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(4 + self.rgdb.len() * 2);
        buf.extend_from_slice(&self.db_rtrw.to_le_bytes());
        for offset in &self.rgdb {
            buf.extend_from_slice(&offset.to_le_bytes());
        }
        buf
    }
}
