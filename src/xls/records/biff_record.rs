// src/biff_records/mod.rs

pub const MAX_RECORD_DATA_SIZE: usize = 8224; // BIFF limit per record
pub const CONTINUE_RECORD_ID: u16 = 0x003C;

pub trait BiffRecord {
    fn id(&self) -> u16;
    fn data(&self) -> Vec<u8>;

    /// Serialize this record, including auto-generation of CONTINUE records if needed
    fn serialize(&self) -> Vec<u8> {
        let payload = self.data();

        // Split into multiple records using CONTINUE
        let first_chunk_size = MAX_RECORD_DATA_SIZE.min(payload.len());
        let (first_chunk, mut remaining_data) = payload.split_at(first_chunk_size);

        let mut result = {
            let len = first_chunk.len() as u16;
            let mut buf = Vec::with_capacity(4 + first_chunk.len());
            buf.extend_from_slice(&self.id().to_le_bytes());
            buf.extend_from_slice(&len.to_le_bytes());
            buf.extend_from_slice(first_chunk);
            buf
        };

        // 生成 CONTINUE records 处理剩余数据
        while !remaining_data.is_empty() {
            let chunk_size = MAX_RECORD_DATA_SIZE.min(remaining_data.len());
            let (chunk, rest) = remaining_data.split_at(chunk_size);

            let len = chunk.len() as u16;
            result.extend_from_slice(&CONTINUE_RECORD_ID.to_le_bytes());
            result.extend_from_slice(&len.to_le_bytes());
            result.extend_from_slice(chunk);

            remaining_data = rest;
        }

        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    struct TestLargeRecord {
        data: Vec<u8>,
    }

    impl TestLargeRecord {
        fn new(size: usize) -> Self {
            Self {
                data: vec![0x41; size], // 填充 'A'
            }
        }
    }

    impl BiffRecord for TestLargeRecord {
        fn id(&self) -> u16 {
            0x1234
        }
        fn data(&self) -> Vec<u8> {
            self.data.clone()
        }
    }

    #[test]
    fn test_continue_splitting() {
        let large_record = TestLargeRecord::new(10000);
        let serialized = large_record.serialize();

        // 验证至少有两个 record（原始 + 至少一个 CONTINUE）
        assert!(serialized.len() > 8224 + 4);

        // 验证包含 CONTINUE record ID
        let continue_bytes = CONTINUE_RECORD_ID.to_le_bytes();
        assert!(serialized.windows(2).any(|w| w == continue_bytes));
    }
}
