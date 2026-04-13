// src/biff_records/mod.rs

/// BIFF记录的最大数据大小限制（8224字节）
/// 超过此大小的数据需要使用CONTINUE记录分片
pub const MAX_RECORD_DATA_SIZE: usize = 8224;
/// CONTINUE记录的记录ID
pub const CONTINUE_RECORD_ID: u16 = 0x003C;

/// BIFF记录trait，用于Excel BIFF8格式记录的序列化和反序列化
/// 所有BIFF记录都应实现此trait
///
/// ## 作用
///
/// Excel文件采用BIFF (Binary Interchange File Format) 格式存储数据。每个BIFF记录由
/// 4字节头部（2字节记录ID + 2字节数据长度）和可变长度的数据部分组成。
///
/// ## 参数说明
///
/// - `id()`: 返回记录的唯一标识符（2字节），用于标识记录类型
/// - `data()`: 返回记录的数据部分（不含头部）
/// - `serialize()`: 自动序列化记录，包含头部和必要的CONTINUE分片
///
/// ## 实现说明
///
/// 大多数记录只需实现 `id()` 和 `data()` 方法，`serialize()` 提供了默认实现，
/// 会自动处理超过8224字节的大数据分片问题。某些记录（如MergedCellsRecord）
/// 可以重写 `serialize()` 方法返回空数据以跳过空记录。
pub trait BiffRecord {
    /// 返回记录的唯一标识符
    fn id(&self) -> u16;
    /// 返回记录的原始数据（不含头部）
    fn data(&self) -> Vec<u8>;

    /// Serialize this record, including auto-generation of CONTINUE records if needed
    /// 自动将超过MAX_RECORD_DATA_SIZE的数据分割为多个CONTINUE记录
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

    #[test]
    fn test_serialize_simple() {
        println!("test_serialize_simple: starting");
        let record = TestLargeRecord::new(10);
        println!("test_serialize_simple: created record");
        let result = record.serialize();
        println!(
            "test_serialize_simple: serialize returned {} bytes",
            result.len()
        );
        assert!(result.len() > 0);
    }

    #[test]
    fn test_serialize_with_u16_id() {
        println!("test_serialize_with_u16_id: starting");

        struct SimpleRecord {
            id: u16,
            data: Vec<u8>,
        }

        impl BiffRecord for SimpleRecord {
            fn id(&self) -> u16 {
                self.id
            }
            fn data(&self) -> Vec<u8> {
                self.data.clone()
            }
        }

        let record = SimpleRecord {
            id: 0x0208,
            data: vec![0; 16],
        };

        println!("test_serialize_with_u16_id: created record, calling serialize");
        let result = record.serialize();
        println!(
            "test_serialize_with_u16_id: returned {} bytes",
            result.len()
        );
    }
}
