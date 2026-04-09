use super::BiffRecord;

/// BackupRecord 记录
///
/// 作用：控制是否创建备份文件
///
/// BackupRecord是Excel BIFF格式中的备份记录（ID: 0x0040），用于指示Excel
/// 在保存文件时是否创建备份文件（.xls -> .xls~）。
///
/// ## 参数说明
///
/// - `backup`: 备份标志
///   - false = 不创建备份（默认）
///   - true = 保存前创建备份
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_backup_record_id() {
        let record = BackupRecord::new(false);
        assert_eq!(record.id(), 0x0040);
    }

    #[test]
    fn test_backup_record_default() {
        let record = BackupRecord::default();
        assert_eq!(record.backup, false);
    }

    #[test]
    fn test_backup_record_data_size() {
        let record = BackupRecord::new(false);
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_backup_record_enabled() {
        let record = BackupRecord::new(true);
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }

    #[test]
    fn test_backup_record_disabled() {
        let record = BackupRecord::new(false);
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }
}
