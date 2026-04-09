use super::BiffRecord;

#[derive(Debug, Clone)]
pub struct BackupRecord {
    backup: bool,
}

impl BackupRecord {
    pub fn new(backup: bool) -> Self {
        Self { backup }
    }
}

impl BiffRecord for BackupRecord {
    fn id(&self) -> u16 {
        0x0040
    }

    fn data(&self) -> Vec<u8> {
        if self.backup {
            vec![0x01, 0x00]
        } else {
            vec![0x00, 0x00]
        }
    }
}

impl Default for BackupRecord {
    fn default() -> Self {
        Self::new(false)
    }
}
