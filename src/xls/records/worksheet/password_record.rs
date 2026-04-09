use crate::xls::records::BiffRecord;

/// 计算Excel密码的哈希值
///
/// ## 作用
///
/// 根据明文密码计算Excel BIFF格式使用的密码哈希值。
///
/// ## 参数说明
///
/// - `plaintext`: 明文密码字符串
/// - 返回值：16位密码哈希值
///
/// ## 说明
///
/// Excel密码使用特定的哈希算法，与标准加密不同，只是简单的混淆。
/// 空密码的哈希值为0。
pub fn password_hash(plaintext: &str) -> u16 {
    if plaintext.is_empty() {
        return 0;
    }

    let mut passwd_hash: u16 = 0;
    for (i, c) in plaintext.char_indices() {
        let c = (c as u32) << (i + 1);
        let low_15 = c & 0x7fff;
        let high_15 = ((c & 0x7fff) << 15) >> 15;
        let c = low_15 | high_15;
        passwd_hash ^= c as u16;
    }
    passwd_hash ^= plaintext.len() as u16;
    passwd_hash ^= 0xCE4B;
    passwd_hash
}

/// PasswordRecord 记录（工作表密码）
///
/// 作用：存储工作表的密码哈希值
///
/// PasswordRecord是Excel BIFF格式中的工作表密码记录（ID: 0x0013），用于
/// 为工作表设置密码保护。密码以哈希值形式存储，而非明文。
///
/// ## 参数说明
///
/// - `passwd`: 明文密码，构造函数会自动计算哈希值
/// - `value`: 密码哈希值（0表示无密码）
#[derive(Debug)]
pub struct PasswordRecord {
    value: u16,
}

impl PasswordRecord {
    pub fn new(passwd: &str) -> Self {
        PasswordRecord {
            value: password_hash(passwd),
        }
    }
}

impl Default for PasswordRecord {
    fn default() -> Self {
        PasswordRecord { value: 0 }
    }
}

impl BiffRecord for PasswordRecord {
    fn id(&self) -> u16 {
        0x0013
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(2);
        buf.extend_from_slice(&self.value.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_password_record_id() {
        let record = PasswordRecord::default();
        assert_eq!(record.id(), 0x0013);
    }

    #[test]
    fn test_password_record_default() {
        let record = PasswordRecord::default();
        assert_eq!(record.value, 0);
    }

    #[test]
    fn test_password_record_data_size() {
        let record = PasswordRecord::default();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_password_hash_empty() {
        assert_eq!(password_hash(""), 0);
    }

    #[test]
    fn test_password_hash_simple() {
        let hash = password_hash("test");
        assert_ne!(hash, 0);
    }
}
