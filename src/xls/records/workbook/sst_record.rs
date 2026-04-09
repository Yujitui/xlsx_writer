use super::encode_biff_string_v2;
use super::BiffRecord;
use std::collections::HashMap;

/// SST记录的记录ID
const SST_RECORD_ID: u16 = 0x00FC;
/// CONTINUE记录的记录ID
const CONTINUE_RECORD_ID: u16 = 0x003C;
/// BIFF记录的最大数据大小
const MAX_RECORD_DATA_SIZE: usize = 8224;

/// 共享字符串表 (Shared String Table)
/// 存储工作簿中所有唯一的字符串，用于LabelSSTRecord引用
#[derive(Debug)]
pub struct SharedStringTable {
    strings: Vec<String>,
    index_map: HashMap<String, usize>,
    counts: Vec<usize>,
    total_refs: usize,
}

impl SharedStringTable {
    pub fn new() -> Self {
        SharedStringTable {
            strings: Vec::new(),
            index_map: HashMap::new(),
            counts: Vec::new(),
            total_refs: 0,
        }
    }

    pub fn add(&mut self, s: String) -> usize {
        self.total_refs += 1;

        if let Some(&idx) = self.index_map.get(&s) {
            self.counts[idx] += 1;
            return idx;
        }

        let idx = self.strings.len();
        self.index_map.insert(s.clone(), idx);
        self.strings.push(s);
        self.counts.push(1);
        idx
    }

    pub fn string_count(&self) -> usize {
        self.strings.len()
    }

    pub fn total_reference_count(&self) -> usize {
        self.total_refs
    }

    pub fn get_strings(&self) -> &[String] {
        &self.strings
    }

    pub fn unique_count(&self) -> usize {
        self.strings.len()
    }
}

impl Default for SharedStringTable {
    fn default() -> Self {
        Self::new()
    }
}

pub struct SSTRecord {
    total_refs: usize,
    unique_count: usize,
    strings: Vec<String>,
}

impl SSTRecord {
    pub fn from(table: &SharedStringTable) -> Self {
        SSTRecord {
            total_refs: table.total_reference_count(),
            unique_count: table.string_count(),
            strings: table.get_strings().to_vec(),
        }
    }

    fn encode_all_strings(&self) -> Vec<u8> {
        let mut data = Vec::new();
        for s in &self.strings {
            data.extend_from_slice(&encode_biff_string_v2(s));
        }
        data
    }
}

impl BiffRecord for SSTRecord {
    fn id(&self) -> u16 {
        SST_RECORD_ID
    }

    fn data(&self) -> Vec<u8> {
        let mut data = Vec::new();

        // 4 bytes: total reference count
        data.extend_from_slice(&(self.total_refs as u32).to_le_bytes());

        // 4 bytes: unique string count
        data.extend_from_slice(&(self.unique_count as u32).to_le_bytes());

        // strings
        data.extend_from_slice(&self.encode_all_strings());

        data
    }

    fn serialize(&self) -> Vec<u8> {
        let encoded = self.encode_all_strings();
        let data_size = 8 + encoded.len(); // 8 bytes header + strings

        // 如果数据小于等于最大容量，直接返回（添加 ID 和长度头）
        if data_size <= MAX_RECORD_DATA_SIZE {
            let mut result = Vec::new();
            result.extend_from_slice(&SST_RECORD_ID.to_le_bytes());
            result.extend_from_slice(&(data_size as u16).to_le_bytes());
            result.extend_from_slice(&self.data());
            return result;
        }

        // 需要分片
        let mut result = Vec::new();
        let mut remaining = &encoded[..];
        let max_data_size = MAX_RECORD_DATA_SIZE - 8; // 减去头部

        // 第一个 SST 记录
        let first_chunk_size = max_data_size.min(encoded.len());
        let (first, rest) = remaining.split_at(first_chunk_size);

        // SST 头部
        result.extend_from_slice(&SST_RECORD_ID.to_le_bytes());
        let sst_data_len = (8 + first.len()) as u16;
        result.extend_from_slice(&sst_data_len.to_le_bytes());
        result.extend_from_slice(&(self.total_refs as u32).to_le_bytes());
        result.extend_from_slice(&(self.unique_count as u32).to_le_bytes());
        result.extend_from_slice(first);

        remaining = rest;

        // 后续 CONTINUE 记录
        while !remaining.is_empty() {
            let chunk_size = max_data_size.min(remaining.len());
            let (chunk, rest) = remaining.split_at(chunk_size);

            // CONTINUE 头部
            result.extend_from_slice(&CONTINUE_RECORD_ID.to_le_bytes());
            result.extend_from_slice(&(chunk.len() as u16).to_le_bytes());
            result.extend_from_slice(chunk);

            remaining = rest;
        }

        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sst_record_id() {
        let mut table = SharedStringTable::new();
        table.add("Hello".to_string());
        let record = SSTRecord::from(&table);
        assert_eq!(record.id(), 0x00FC);
    }

    #[test]
    fn test_sst_add_string() {
        let mut table = SharedStringTable::new();
        let idx1 = table.add("Hello".to_string());
        let idx2 = table.add("World".to_string());
        let idx3 = table.add("Hello".to_string());

        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(idx3, 0); // 返回已存在字符串的索引
        assert_eq!(table.string_count(), 2);
        assert_eq!(table.total_reference_count(), 3);
    }

    #[test]
    fn test_sst_record_data() {
        let mut table = SharedStringTable::new();
        table.add("Hello".to_string());
        table.add("World".to_string());
        let record = SSTRecord::from(&table);
        let data = record.data();

        // Check header
        assert_eq!(&data[0..4], &2u32.to_le_bytes()); // total_refs
        assert_eq!(&data[4..8], &2u32.to_le_bytes()); // unique_count
    }

    #[test]
    fn test_sst_record_serialize_no_continue() {
        let mut table = SharedStringTable::new();
        table.add("Test".to_string());
        let record = SSTRecord::from(&table);
        let serialized = record.serialize();

        // SST ID
        assert_eq!(&serialized[0..2], &SST_RECORD_ID.to_le_bytes());
    }

    #[test]
    fn test_sst_record_with_unicode() {
        let mut table = SharedStringTable::new();
        table.add("测试".to_string());
        let record = SSTRecord::from(&table);
        let data = record.data();

        // Check header
        assert_eq!(&data[0..4], &1u32.to_le_bytes());
        assert_eq!(&data[4..8], &1u32.to_le_bytes());
    }
}
