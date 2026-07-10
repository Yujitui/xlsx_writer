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
        // BIFF8 WriteAccess: u16 cch + u8 fHighByte + cch bytes of short string + spaces padding to 112 bytes
        let name_bytes = self.username.as_bytes();
        let copy_len = std::cmp::min(name_bytes.len(), 109);
        let mut buf = Vec::with_capacity(112);
        buf.extend_from_slice(&(copy_len as u16).to_le_bytes()); // cch
        buf.push(0x00); // fHighByte (compressed)
        buf.extend_from_slice(&name_bytes[..copy_len]);
        buf.resize(112, 0x20); // pad with spaces
        buf
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
        assert_eq!(&data[0..2], &3u16.to_le_bytes());
        assert_eq!(data[2], 0x00);
        assert_eq!(&data[3..6], b"ABC");
    }
}
