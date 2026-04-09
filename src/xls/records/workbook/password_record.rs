use super::BiffRecord;
use std::fmt;

#[derive(Debug, Clone)]
pub struct PasswordRecord {
    password_hash: u16,
}

impl PasswordRecord {
    /// 创建新的密码记录
    pub fn new(password: &str) -> Self {
        Self {
            password_hash: Self::hash_password(password),
        }
    }

    /// 基于OpenOffice算法的密码哈希实现
    fn hash_password(password: &str) -> u16 {
        if password.is_empty() {
            return 0;
        }

        let mut hash: u16 = 0x0000;
        for (i, ch) in password.chars().enumerate() {
            let char_code = ch as u16;
            let shifted = char_code << (i + 1);
            let low_15 = shifted & 0x7FFF;
            let high_15 = (shifted & 0x7FFFu16.wrapping_shl(15)) >> 15;
            let combined = low_15 | high_15;
            hash ^= combined;
        }
        hash ^= password.len() as u16;
        hash ^= 0xCE4B;
        hash
    }

    /// 获取密码哈希值
    pub fn get_hash(&self) -> u16 {
        self.password_hash
    }

    /// 检查密码是否匹配（通过哈希比较）
    pub fn verify_password(&self, password: &str) -> bool {
        self.password_hash == Self::hash_password(password)
    }

    /// 创建空密码记录
    pub fn empty() -> Self {
        Self { password_hash: 0 }
    }
}

impl BiffRecord for PasswordRecord {
    fn id(&self) -> u16 {
        0x0013 // Password record ID
    }

    fn data(&self) -> Vec<u8> {
        self.password_hash.to_le_bytes().to_vec()
    }
}

impl Default for PasswordRecord {
    fn default() -> Self {
        Self::empty()
    }
}

impl fmt::Display for PasswordRecord {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "PasswordHash({:04X})", self.password_hash)
    }
}
