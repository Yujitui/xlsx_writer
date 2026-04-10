use super::encode_biff_string_v2;
use super::BiffRecord;
use crate::xls::XlsError;
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

    /// 添加字符串到 SST（用于解析时，不检查重复）
    pub fn push_string(&mut self, s: String) {
        let idx = self.strings.len();
        self.index_map.insert(s.clone(), idx);
        self.strings.push(s);
        self.counts.push(1);
    }

    /// 从 SST 记录数据解析（包含可能的 CONTINUE 记录）
    ///
    /// # 参数
    /// - `data`: SST 记录的数据部分（已去除 record header）
    /// - `next_record_id`: 查看下一个记录的 ID（用于检测 CONTINUE）
    /// - `read_continue`: 回调函数，如果下一个是 CONTINUE 则读取并返回其数据
    pub fn parse_with_continue<F>(
        data: &[u8],
        mut next_record_id: F,
        mut read_continue: F,
    ) -> Result<Self, XlsError>
    where
        F: FnMut() -> Result<Option<(u16, Vec<u8>)>, XlsError>,
    {
        if data.len() < 8 {
            return Err(XlsError::InvalidFormat(format!(
                "SST data too short: {} bytes",
                data.len()
            )));
        }

        // 读取头部
        let total_refs = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
        let unique_count = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;

        let mut table = SharedStringTable::new();
        table.total_refs = total_refs as usize;

        // 收集所有数据（包括 CONTINUE）
        let mut all_data = data[8..].to_vec();

        // 检查并读取 CONTINUE 记录
        loop {
            match next_record_id()? {
                Some((id, _)) if id == CONTINUE_RECORD_ID => {
                    if let Some((_, cont_data)) = read_continue()? {
                        all_data.extend_from_slice(&cont_data);
                    } else {
                        break;
                    }
                }
                _ => break,
            }
        }

        // 解析字符串
        let mut offset = 0;
        for _ in 0..unique_count {
            if offset >= all_data.len() {
                return Err(XlsError::InvalidFormat(
                    "SST data incomplete while parsing strings".to_string(),
                ));
            }

            // 读取字符数（2字节）
            if offset + 2 > all_data.len() {
                return Err(XlsError::InvalidFormat(
                    "SST string header incomplete".to_string(),
                ));
            }
            let char_count = u16::from_le_bytes([all_data[offset], all_data[offset + 1]]) as usize;
            offset += 2;

            // 读取标志字节
            if offset >= all_data.len() {
                return Err(XlsError::InvalidFormat(
                    "SST string flag byte missing".to_string(),
                ));
            }
            let flag = all_data[offset];
            offset += 1;

            let is_utf16 = (flag & 0x01) != 0;
            let has_rich = (flag & 0x08) != 0;
            let has_ext = (flag & 0x04) != 0;

            // 读取 Rich Text 信息
            let mut rich_runs = 0u16;
            if has_rich {
                if offset + 2 > all_data.len() {
                    return Err(XlsError::InvalidFormat(
                        "SST rich text info incomplete".to_string(),
                    ));
                }
                rich_runs = u16::from_le_bytes([all_data[offset], all_data[offset + 1]]);
                offset += 2;
            }

            // 读取 Extension 信息
            let mut ext_size = 0u32;
            if has_ext {
                if offset + 4 > all_data.len() {
                    return Err(XlsError::InvalidFormat(
                        "SST extension info incomplete".to_string(),
                    ));
                }
                ext_size = u32::from_le_bytes([
                    all_data[offset],
                    all_data[offset + 1],
                    all_data[offset + 2],
                    all_data[offset + 3],
                ]);
                offset += 4;
            }

            // 读取字符串数据
            let string_bytes = if is_utf16 { char_count * 2 } else { char_count };

            if offset + string_bytes > all_data.len() {
                return Err(XlsError::InvalidFormat(format!(
                    "SST string data incomplete: need {} bytes, have {}",
                    string_bytes,
                    all_data.len() - offset
                )));
            }

            let string = if is_utf16 {
                // UTF-16LE 解码
                let utf16_data = &all_data[offset..offset + string_bytes];
                let u16_vec: Vec<u16> = utf16_data
                    .chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                    .collect();
                String::from_utf16(&u16_vec)
                    .unwrap_or_else(|_| String::from_utf8_lossy(utf16_data).to_string())
            } else {
                // ASCII/Latin-1 解码
                let ascii_data = &all_data[offset..offset + string_bytes];
                String::from_utf8_lossy(ascii_data).to_string()
            };
            offset += string_bytes;

            // 添加到表
            table.strings.push(string);
            table.counts.push(1);

            // 跳过 Rich Text 和 Extension 数据
            let skip_bytes = (rich_runs as usize * 4) + ext_size as usize;
            offset += skip_bytes;
        }

        // 重建 index_map
        for (idx, s) in table.strings.iter().enumerate() {
            table.index_map.insert(s.clone(), idx);
        }

        Ok(table)
    }

    /// 从合并的数据解析（SST + CONTINUE 数据已合并）
    pub fn parse_from_data(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 8 {
            return Err(XlsError::InvalidFormat(format!(
                "SST data too short: {} bytes",
                data.len()
            )));
        }

        // 读取头部
        let _total_refs = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
        let unique_count = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;

        let mut table = SharedStringTable::new();
        table.total_refs = _total_refs as usize;

        // 解析字符串
        let mut offset = 8;
        for _ in 0..unique_count {
            if offset >= data.len() {
                break; // 数据不完整，但继续处理已解析的字符串
            }

            // 读取字符数（2字节）
            if offset + 2 > data.len() {
                break;
            }
            let char_count = u16::from_le_bytes([data[offset], data[offset + 1]]) as usize;
            offset += 2;

            // 读取标志字节
            if offset >= data.len() {
                break;
            }
            let flag = data[offset];
            offset += 1;

            let is_utf16 = (flag & 0x01) != 0;
            let has_rich = (flag & 0x08) != 0;
            let has_ext = (flag & 0x04) != 0;

            // 读取 Rich Text 信息
            if has_rich && offset + 2 <= data.len() {
                offset += 2;
            }

            // 读取 Extension 信息
            if has_ext && offset + 4 <= data.len() {
                offset += 4;
            }

            // 读取字符串数据
            let string_bytes = if is_utf16 { char_count * 2 } else { char_count };

            if offset + string_bytes > data.len() {
                break;
            }

            let string = if is_utf16 {
                let utf16_data = &data[offset..offset + string_bytes];
                let u16_vec: Vec<u16> = utf16_data
                    .chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                    .collect();
                String::from_utf16(&u16_vec).unwrap_or_default()
            } else {
                String::from_utf8_lossy(&data[offset..offset + string_bytes]).to_string()
            };
            offset += string_bytes;

            // 添加到表
            table.strings.push(string);
            table.counts.push(1);
        }

        // 重建 index_map
        for (idx, s) in table.strings.iter().enumerate() {
            table.index_map.insert(s.clone(), idx);
        }

        Ok(table)
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

// ============================================================================
// ParsableRecord implementations for reading
// ============================================================================

use crate::xls::records::parseable::{ParsableRecord, ParseState};

/// SST 记录包装结构（用于解析）
pub struct SSTRecordData {
    data: Vec<u8>,
}

impl ParsableRecord for SSTRecordData {
    const RECORD_ID: u16 = 0x00FC;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        Ok(SSTRecordData {
            data: data.to_vec(),
        })
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        // 完成之前的 SST
        if let Some(parser) = state.sst_parser.take() {
            if let Err(e) = parser.finish(&mut state.sst) {
                eprintln!("Warning: Failed to finish previous SST: {}", e);
            }
        }

        // 解析头部，开始新的 SST
        if self.data.len() < 8 {
            return Err(XlsError::InvalidFormat("SST data too short".to_string()));
        }

        let total_refs =
            u32::from_le_bytes([self.data[0], self.data[1], self.data[2], self.data[3]]);
        let unique_count =
            u32::from_le_bytes([self.data[4], self.data[5], self.data[6], self.data[7]]);

        let mut parser = crate::xls::records::SSTParserState::new(total_refs, unique_count);

        // 解析第一个数据块
        if let Err(e) = parser.parse_chunk(&self.data[8..], &mut state.sst) {
            eprintln!("Warning: SST parse error: {}", e);
        }

        state.sst_parser = Some(parser);
        Ok(())
    }
}

/// CONTINUE 记录
pub struct ContinueRecord {
    data: Vec<u8>,
}

impl ParsableRecord for ContinueRecord {
    const RECORD_ID: u16 = 0x003C;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        Ok(ContinueRecord {
            data: data.to_vec(),
        })
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        match state.sst_parser.as_mut() {
            Some(parser) => {
                if let Err(e) = parser.parse_chunk(&self.data, &mut state.sst) {
                    eprintln!("Warning: CONTINUE parse error: {}", e);
                }
            }
            None => {
                eprintln!("Warning: CONTINUE without SST");
            }
        }
        Ok(())
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
