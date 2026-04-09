use super::BiffRecord;

#[derive(Debug)]
pub struct ProtectRecord {
    is_protected: bool,
}

impl ProtectRecord {
    pub fn new(is_protected: bool) -> Self {
        Self { is_protected }
    }

    pub fn protected() -> Self {
        Self::new(true)
    }

    pub fn unprotected() -> Self {
        Self::new(false)
    }
}

impl Default for ProtectRecord {
    fn default() -> Self {
        Self::unprotected()
    }
}

impl BiffRecord for ProtectRecord {
    fn id(&self) -> u16 {
        0x0012 // Protect record ID
    }

    fn data(&self) -> Vec<u8> {
        if self.is_protected {
            vec![0x01, 0x00] // 1 = Protected
        } else {
            vec![0x00, 0x00] // 0 = Not protected
        }
    }
}
