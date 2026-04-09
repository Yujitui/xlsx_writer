use super::BiffRecord;

/// WriteAccessRecord 记录
///
/// 作用：存储文件作者的名称
///
/// WriteAccessRecord是Excel BIFF格式中的写入访问记录（ID: 0x005C），用于
/// 存储创建或最后修改该文件的用户名称。固定112字节宽度。
///
/// ## 参数说明
///
/// - `username`: 用户名（最长111字符，超出部分截断，不足部分用空格填充）
/// - 默认值："None"
#[derive(Debug)]
pub struct WriteAccessRecord {
    username: String,
}

impl WriteAccessRecord {
    pub fn new(username: &str) -> Self {
        Self {
            username: username.to_string(),
        }
    }
}

impl Default for WriteAccessRecord {
    fn default() -> Self {
        Self::new("None")
    }
}

impl BiffRecord for WriteAccessRecord {
    fn id(&self) -> u16 {
        0x005C // Write Access record ID
    }

    fn data(&self) -> Vec<u8> {
        // Username is stored as a fixed-length string (always 112 bytes in BIFF8)
        let mut username_bytes = vec![0u8; 112];

        // Convert username to bytes and copy into the fixed buffer
        let name_bytes = self.username.as_bytes();
        let copy_len = std::cmp::min(name_bytes.len(), 112);
        username_bytes[..copy_len].copy_from_slice(&name_bytes[..copy_len]);

        // Fill remaining bytes with spaces (0x20)
        for i in copy_len..112 {
            username_bytes[i] = 0x20; // Space character
        }

        username_bytes
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_write_access_record_id() {
        let record = WriteAccessRecord::new("Test");
        assert_eq!(record.id(), 0x005C);
    }

    #[test]
    fn test_write_access_record_default() {
        let record = WriteAccessRecord::default();
        assert_eq!(record.username, "None");
    }

    #[test]
    fn test_write_access_record_data_size() {
        let record = WriteAccessRecord::new("Test");
        assert_eq!(record.data().len(), 112);
    }

    #[test]
    fn test_write_access_record_username_encoding() {
        let record = WriteAccessRecord::new("ABC");
        let data = record.data();
        assert_eq!(data[0], b'A');
        assert_eq!(data[1], b'B');
        assert_eq!(data[2], b'C');
    }
}
