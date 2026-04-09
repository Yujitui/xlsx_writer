use super::BiffRecord;

#[repr(u16)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BofType {
    WorkbookGlobals = 0x0005,
    VisualBasicModule = 0x0006,
    Worksheet = 0x0010,
    Chart = 0x0020,
    MacroSheet = 0x0040,
    Workspace = 0x0100,
}

impl BofType {
    pub fn to_u16(&self) -> u16 {
        *self as u16
    }
}

#[derive(Debug)]
pub struct BoFRecord {
    pub bof_type: BofType, // Type of BOF record
}

impl BoFRecord {
    pub fn new(bof_type: BofType) -> Self {
        BoFRecord { bof_type }
    }
}

impl BiffRecord for BoFRecord {
    fn id(&self) -> u16 {
        0x0809 // BOF record ID
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(10);
        buf.extend_from_slice(&0x0600u16.to_le_bytes()); // Version
        buf.extend_from_slice(&(self.bof_type as u16).to_le_bytes()); // Type
        buf.extend_from_slice(&0x0DBBu16.to_le_bytes()); // Build
        buf.extend_from_slice(&0x07CCu16.to_le_bytes()); // Year
        buf.extend_from_slice(&0x00u8.to_le_bytes()); // Flags
        buf.extend_from_slice(&0x06u8.to_le_bytes()); // VerCanRead
        buf
    }
}
