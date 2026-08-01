use super::BiffRecord;
use crate::XlsError;
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

        // 使用规范的状态机逐块解析，正确处理 CONTINUE 续接标志字节
        let mut parser = crate::xls_records::SSTParserState::new(total_refs, unique_count as u32);
        parser.parse_chunk(&data[8..], &mut table)?;

        loop {
            match next_record_id()? {
                Some((id, _)) if id == CONTINUE_RECORD_ID => {
                    if let Some((_, cont_data)) = read_continue()? {
                        parser.parse_chunk(&cont_data, &mut table)?;
                    } else {
                        break;
                    }
                }
                _ => break,
            }
        }

        parser.finish(&mut table)?;

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

/// 编码 SST 中的字符串，匹配 Excel 参考文件格式。
///
/// Excel 为 SST 字符串添加 Asian phonetic extension（16 字节），
/// 即使字符串本身不包含亚洲语音信息。flag 的 bit 2 表示扩展存在。
fn encode_sst_string(s: &str) -> Vec<u8> {
    let mut result = Vec::new();

    if s.chars().all(|c| c <= '\u{00FF}') {
        let bytes: Vec<u8> = s.chars().map(|c| c as u8).collect();
        let char_count = bytes.len() as u16;
        result.extend_from_slice(&char_count.to_le_bytes());
        result.push(0x04); // bit 2 = Asian phonetic extension present
        result.extend_from_slice(&16u32.to_le_bytes()); // cbExtRst
        result.extend_from_slice(&bytes);
    } else {
        let utf16: Vec<u8> = s.encode_utf16().flat_map(|c| c.to_le_bytes()).collect();
        let char_count = (utf16.len() / 2) as u16;
        result.extend_from_slice(&char_count.to_le_bytes());
        result.push(0x05); // bit 0 = unicode, bit 2 = extension present
        result.extend_from_slice(&16u32.to_le_bytes()); // cbExtRst
        result.extend_from_slice(&utf16);
    }

    // 16 bytes Asian phonetic extension data（与参考文件一致）
    result.extend_from_slice(&[
        0x01, 0x00, 0x0c, 0x00, 0x06, 0x00, 0x37, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    ]);

    result
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
            data.extend_from_slice(&encode_sst_string(s));
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

use crate::xls_records::parseable::{ParsableRecord, ParseState};

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

        let mut parser = crate::xls_records::SSTParserState::new(total_refs, unique_count);

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
        assert_eq!(idx3, 0);
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

        assert_eq!(&data[0..4], &2u32.to_le_bytes());
        assert_eq!(&data[4..8], &2u32.to_le_bytes());
    }

    #[test]
    fn test_sst_record_serialize_no_continue() {
        let mut table = SharedStringTable::new();
        table.add("Test".to_string());
        let record = SSTRecord::from(&table);
        let serialized = record.serialize();

        assert_eq!(&serialized[0..2], &SST_RECORD_ID.to_le_bytes());
    }

    #[test]
    fn test_sst_record_with_unicode() {
        let mut table = SharedStringTable::new();
        table.add("测试".to_string());
        let record = SSTRecord::from(&table);
        let data = record.data();

        assert_eq!(&data[0..4], &1u32.to_le_bytes());
        assert_eq!(&data[4..8], &1u32.to_le_bytes());
    }

    #[test]
    fn test_sst_record_data_matches_excel_reference() {
        let mut table = SharedStringTable::new();
        table.add("aaa".to_string());
        table.add("bbb".to_string());
        let record = SSTRecord::from(&table);
        let data = record.data();

        let expected = vec![
            0x02, 0x00, 0x00, 0x00,
            0x02, 0x00, 0x00, 0x00,
            0x03, 0x00, 0x04,
            0x10, 0x00, 0x00, 0x00,
            0x61, 0x61, 0x61,
            0x01, 0x00, 0x0c, 0x00, 0x06, 0x00, 0x37, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x03, 0x00, 0x04,
            0x10, 0x00, 0x00, 0x00,
            0x62, 0x62, 0x62,
            0x01, 0x00, 0x0c, 0x00, 0x06, 0x00, 0x37, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];
        assert_eq!(data, expected, "SST 数据与 Excel 参考不匹配");
    }

    #[test]
    fn test_sst_string_unicode_has_extension() {
        let encoded = encode_sst_string("测试");
        assert_eq!(&encoded[0..2], &2u16.to_le_bytes());
        assert_eq!(encoded[2], 0x05);
        assert_eq!(&encoded[3..7], &16u32.to_le_bytes());
        assert_eq!(&encoded[7..11], &[0x4B, 0x6D, 0xD5, 0x8B]);
        assert_eq!(encoded.len(), 11 + 16);
    }

    // ========================================================================
    // 长度一致性验证: encode_sst_string 各部分之和等于总长度
    // ========================================================================

    /// 验证单个字符串编码后的长度正确性
    fn verify_encoded_string_length(s: &str) {
        let encoded = encode_sst_string(s);
        // 结构: cch(2) + flag(1) + cbExtRst(4) + chars(N or 2*N) + extData(16)
        let expected_body = if s.chars().all(|c| c <= '\u{00FF}') {
            s.len()
        } else {
            s.encode_utf16().count() * 2
        };
        let expected_total = 2 + 1 + 4 + expected_body + 16;
        assert_eq!(
            encoded.len(),
            expected_total,
            "encode_sst_string({:?}) length mismatch: expected {} + {}, got {}",
            s, expected_body, 23, encoded.len()
        );
    }

    #[test]
    fn test_encode_sst_string_length_ascii() {
        for s in &["", "a", "abc", "hello world", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"] {
            verify_encoded_string_length(s);
        }
    }

    #[test]
    fn test_encode_sst_string_length_unicode() {
        for s in &["测", "测试", "中文", "日本語", "中文English混合", "🌍"] {
            verify_encoded_string_length(s);
        }
    }

    #[test]
    fn test_encode_sst_string_length_mixed() {
        // 混合 ASCII + Unicode
        let tests = vec![
            "Hello测试",
            "abc中文def",
            "123测试456",
            "a",
            "测",
            "",
        ];
        for s in &tests {
            verify_encoded_string_length(s);
        }
    }

    /// 验证 SST 记录的 data() 总长度 = 8 + sum(每个字符串的编码长度)
    #[test]
    fn test_sst_data_total_length() {
        let strings = vec![
            "Hello", "World", "测试", "中文SST", "abc123", "ABC",
        ];
        let mut table = SharedStringTable::new();
        for s in &strings {
            table.add(s.to_string());
        }
        let record = SSTRecord::from(&table);
        let data = record.data();

        // 计算预期长度: 8(header) + sum(encode_sst_string)
        let expected_encoded_len: usize = strings.iter().map(|s| encode_sst_string(s).len()).sum();
        let expected_total = 8 + expected_encoded_len;

        assert_eq!(
            data.len(),
            expected_total,
            "SST data() total length mismatch: expected {} (8 header + {} strings), got {}",
            expected_total, expected_encoded_len, data.len()
        );
    }

    // ========================================================================
    // 序列化一致性验证: serialize() 的 ID + length + data 自洽
    // ========================================================================

    #[test]
    fn test_sst_serialize_self_consistent() {
        let strings = vec!["Hello", "World", "测试", "中文SST"];
        let mut table = SharedStringTable::new();
        for s in &strings {
            table.add(s.to_string());
        }
        let record = SSTRecord::from(&table);
        let serialized = record.serialize();

        let declared_len = u16::from_le_bytes([serialized[2], serialized[3]]);
        let actual_data_len = serialized.len() - 4;

        // 诊断：对比 encode_all_strings 和 data()
        let encoded = record.encode_all_strings();
        let data = record.data();
        eprintln!(
            "DEBUG: serialized.len={}, declared_len={}, actual_data_len={}, data.len()={}, encoded.len()={}, 8+encoded={}",
            serialized.len(), declared_len, actual_data_len, data.len(), encoded.len(), 8 + encoded.len()
        );

        // 检查 data() = 8 + encode_all_strings() 是否成立
        assert_eq!(
            data.len(),
            8 + encoded.len(),
            "data() length mismatch: data()={} != 8+encode_all()={}",
            data.len(),
            8 + encoded.len()
        );

        assert_eq!(
            declared_len as usize,
            actual_data_len,
            "SST serialized length mismatch: declared {} vs actual data {} (data()={}, encoded={})",
            declared_len, actual_data_len, data.len(), encoded.len()
        );
    }

    // ========================================================================
    // 编码-解析 往返测试（Round-Trip）
    // ========================================================================

    /// 手动解析单条 encode_sst_string 编码的字符串
    fn roundtrip_one_encoded(encoded: &[u8], original: &str) {
        let cch = u16::from_le_bytes([encoded[0], encoded[1]]) as usize;
        let flag = encoded[2];
        let has_unicode = (flag & 0x01) != 0;
        let has_ext = (flag & 0x04) != 0;

        assert!(
            has_ext,
            "flag 必须包含 extension bit (0x04): got 0x{:02X}",
            flag
        );

        let ext_size = if has_ext {
            u32::from_le_bytes([encoded[3], encoded[4], encoded[5], encoded[6]]) as usize
        } else {
            0
        };

        let char_bytes = if has_unicode { cch * 2 } else { cch };
        let string_start = 3 + 4; // flag + cbExtRst
        let string_end = string_start + char_bytes;

        let decoded = if has_unicode {
            let u16_vec: Vec<u16> = encoded[string_start..string_end]
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                .collect();
            String::from_utf16(&u16_vec).unwrap()
        } else {
            String::from_utf8_lossy(&encoded[string_start..string_end]).to_string()
        };

        assert_eq!(
            decoded, original,
            "Round-trip mismatch for {:?}", original
        );

        // 验证 extension 数据长度
        let ext_start = string_end;
        let ext_end = ext_start + ext_size;
        assert!(
            ext_end <= encoded.len(),
            "Extension data overflows: ext_start={} ext_size={} encoded_len={}",
            ext_start,
            ext_size,
            encoded.len()
        );
    }

    #[test]
    fn test_sst_string_roundtrip_ascii() {
        for s in &["", "a", "abc", "hello world"] {
            let encoded = encode_sst_string(s);
            roundtrip_one_encoded(&encoded, s);
        }
    }

    #[test]
    fn test_sst_string_roundtrip_unicode() {
        for s in &["测", "测试", "中文", "日本語", "中文English混合"] {
            let encoded = encode_sst_string(s);
            roundtrip_one_encoded(&encoded, s);
        }
    }

    // ========================================================================
    // CONTINUE 分片测试：大量唯一字符串触发分片
    // ========================================================================

    /// 构建足够大的 SST 触发 CONTINUE 记录，验证 serialize() 完整性
    #[test]
    fn test_sst_continue_split_integrity() {
        let mut table = SharedStringTable::new();
        // 每个 ASCII 字符串编码后固定 23+N = ~28-30 字节
        // 需要超过 MAX_RECORD_DATA_SIZE (8224) 的数据
        // 8224 / ~30 ≈ 274 个字符串即可触发 CONTINUE
        for i in 0..300 {
            table.add(format!("unique_string_{:04}", i));
        }

        let record = SSTRecord::from(&table);
        let serialized = record.serialize();

        // 验证第一个记录是 SST (0x00FC)
        assert_eq!(
            u16::from_le_bytes([serialized[0], serialized[1]]),
            SST_RECORD_ID
        );

        // 扫描所有记录验证格式完整性
        let mut offset = 0;
        let mut record_count = 0;
        let mut total_declared_data = 0usize;

        while offset < serialized.len() {
            if offset + 4 > serialized.len() {
                break;
            }
            let id = u16::from_le_bytes([serialized[offset], serialized[offset + 1]]);
            let len = u16::from_le_bytes([serialized[offset + 2], serialized[offset + 3]]) as usize;
            let total_record_size = 4 + len;

            // 验证记录不越界
            assert!(
                offset + total_record_size <= serialized.len(),
                "Record {} overflows buffer: offset={}, declared_len={}, buf_len={}",
                record_count,
                offset,
                len,
                serialized.len()
            );

            // 验证 ID 有效
            assert!(
                id == SST_RECORD_ID || id == CONTINUE_RECORD_ID,
                "Invalid record ID at offset {}: 0x{:04X}",
                offset,
                id
            );

            // 第一条必须是 SST
            if record_count == 0 {
                assert_eq!(id, SST_RECORD_ID, "First record must be SST");
            }

            total_declared_data += len;
            offset += total_record_size;
            record_count += 1;
        }

        assert!(
            record_count > 1,
            "Should have CONTINUE records (>1 record), got {}",
            record_count
        );
        // 验证总数据长度一致
        assert_eq!(offset, serialized.len(), "Record chain length mismatch");
    }

    // ========================================================================
    // 多区域实际场景模拟：验证 SST 跨 region 的字符串索引正确
    // ========================================================================

    #[test]
    fn test_sst_multi_region_strings() {
        let mut sst = SharedStringTable::new();

        // 模拟多区域中的字符串收集过程
        let region1_strings = ["姓名", "年龄", "张三", "28"];
        let region2_strings = ["产品", "销售额", "笔记本电脑", "9999"];

        let mut all_indices = Vec::new();
        for s in region1_strings.iter().chain(region2_strings.iter()) {
            let idx = sst.add(s.to_string());
            all_indices.push(idx);
        }

        // 验证总引用数和独立字符串数
        assert_eq!(sst.total_reference_count(), 8);
        assert_eq!(sst.unique_count(), 8);

        // 验证 SST 序列化后能正确编码所有字符串
        let record = SSTRecord::from(&sst);
        let data = record.data();
        let encoded_strs = &data[8..]; // 跳过 8 字节 header

        // 手动解析每条字符串，验证与原始一致
        let mut offset = 0;
        let all_originals: Vec<&str> = region1_strings
            .iter()
            .chain(region2_strings.iter())
            .copied()
            .collect();

        for (i, original) in all_originals.iter().enumerate() {
            // 每条字符串长度由 encode_sst_string 决定
            let one_encoded = encode_sst_string(original);
            let expected_len = one_encoded.len();

            assert!(
                offset + expected_len <= encoded_strs.len(),
                "String {} overflows: offset={}, len={}, buf={}",
                i,
                offset,
                expected_len,
                encoded_strs.len()
            );

            let slice = &encoded_strs[offset..offset + expected_len];
            assert_eq!(
                slice, &one_encoded[..],
                "String {} ({:?}) encoding mismatch at offset {}",
                i, original, offset
            );

            offset += expected_len;
        }

        assert_eq!(
            offset, encoded_strs.len(),
            "Total encoded strings length mismatch: consumed {} of {}",
            offset, encoded_strs.len()
        );
    }
}
