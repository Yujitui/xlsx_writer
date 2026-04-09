use super::BiffRecord;

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
