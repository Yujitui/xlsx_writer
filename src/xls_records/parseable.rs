//! 可解析记录 trait
//!
//! 为 BIFF 记录提供统一的反序列化接口

use crate::xls_records::types::XlsSheet;
use crate::xls_records::workbook::sst_record::SharedStringTable;
use crate::XlsError;

/// 可解析记录 trait
///
/// 为需要反序列化的 BIFF 记录实现此 trait
pub trait ParsableRecord: Sized {
    /// 记录类型 ID
    const RECORD_ID: u16;

    /// 从字节数据解析记录
    ///
    /// # 参数
    /// - `data`: 记录数据部分（不含 4 字节 header）
    fn parse(data: &[u8]) -> Result<Self, XlsError>;

    /// 将记录应用到解析状态
    ///
    /// # 参数
    /// - `state`: 可变解析状态，包含当前工作表、SST 等
    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError>;
}

/// 单个字符串在 SST 解析过程中的进度阶段
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SstStringStage {
    /// 正在读取字符串头部（cch / grbit / cRun / cbExtRst）
    Header,
    /// 正在读取字符数据
    Chars,
    /// 正在跳过 Rich Text runs 数据（cRun * 4 字节）
    Rich,
    /// 正在跳过扩展（注音）数据（ext_size 字节）
    Ext,
}

/// 跨边界部分字符串状态
#[derive(Debug)]
pub struct PartialString {
    pub stage: SstStringStage,
    pub char_count: usize,
    pub is_utf16: bool,
    pub has_rich: bool,
    pub has_ext: bool,
    pub rich_runs: u16,
    pub ext_size: u32,
    /// 当前 stage 已消耗的字节数
    pub bytes_read: usize,
    /// Header 阶段的头部字节，或 Chars 阶段当前压缩模式的字符字节
    pub buffer: Vec<u8>,
    /// 已完成的字符段（压缩模式, 字节）。当 CONTINUE 续接标志切换压缩模式时，
    /// 前一段的字节按原模式保存在此，最后一段使用 `buffer`。
    pub segments: Vec<(bool, Vec<u8>)>,
}

/// SST 解析状态机
#[derive(Debug)]
pub struct SSTParserState {
    pub total_refs: u32,
    pub unique_count: u32,
    pub strings_parsed: u32,
    pub current_string: Option<PartialString>,
}

impl SSTParserState {
    pub fn new(total_refs: u32, unique_count: u32) -> Self {
        SSTParserState {
            total_refs,
            unique_count,
            strings_parsed: 0,
            current_string: None,
        }
    }

    /// 解析数据块，将完成的字符串添加到 SST
    ///
    /// 依据 BIFF8 规范处理跨记录边界的字符串：
    /// - 字符串字符数据被分片到 CONTINUE 记录时，续接块以 1 字节 grbit 开头，
    ///   其 bit0 重新指定后续字符的压缩模式（8 位或 UTF-16）。
    /// - Rich Text runs 与扩展数据可以跨块，但续接块前没有该标志字节。
    pub fn parse_chunk(
        &mut self,
        data: &[u8],
        sst: &mut SharedStringTable,
    ) -> Result<(), XlsError> {
        let mut offset = 0;

        loop {
            if self.strings_parsed >= self.unique_count || offset >= data.len() {
                return Ok(());
            }

            match self.current_string.take() {
                // 开始解析新字符串
                None => {
                    let mut partial = PartialString {
                        stage: SstStringStage::Header,
                        char_count: 0,
                        is_utf16: false,
                        has_rich: false,
                        has_ext: false,
                        rich_runs: 0,
                        ext_size: 0,
                        bytes_read: 0,
                        buffer: Vec::new(),
                        segments: Vec::new(),
                    };

                    if !self.read_header(data, &mut offset, &mut partial) {
                        self.current_string = Some(partial);
                        return Ok(());
                    }

                    // 头部完整，读取字符数据
                    if !self.read_chars(data, &mut offset, &mut partial, sst) {
                        self.current_string = Some(partial);
                        return Ok(());
                    }

                    // 字符已入表；若还有 Rich/Ext 数据，继续跳过
                    if partial.has_rich || partial.has_ext {
                        self.current_string = Some(partial);
                    }
                }
                // 恢复未完成的字符串
                Some(mut partial) => {
                    match partial.stage {
                        SstStringStage::Header => {
                            if !self.read_header(data, &mut offset, &mut partial) {
                                self.current_string = Some(partial);
                                return Ok(());
                            }
                            if !self.read_chars(data, &mut offset, &mut partial, sst) {
                                self.current_string = Some(partial);
                                return Ok(());
                            }
                            if partial.has_rich || partial.has_ext {
                                self.current_string = Some(partial);
                            }
                        }
                        SstStringStage::Chars => {
                            // 字符串被分片：续接块首字节为新的 grbit（仅 bit0 有效）
                            if offset >= data.len() {
                                self.current_string = Some(partial);
                                return Ok(());
                            }
                            let cont_flag = data[offset];
                            offset += 1;

                            let new_utf16 = (cont_flag & 0x01) != 0;
                            if new_utf16 != partial.is_utf16 {
                                // 压缩模式切换：把已收集的字节按旧模式保存为独立段
                                if !partial.buffer.is_empty() {
                                    partial.segments.push((
                                        partial.is_utf16,
                                        std::mem::take(&mut partial.buffer),
                                    ));
                                }
                                partial.is_utf16 = new_utf16;
                            }

                            if !self.read_chars(data, &mut offset, &mut partial, sst) {
                                self.current_string = Some(partial);
                                return Ok(());
                            }
                            if partial.has_rich || partial.has_ext {
                                self.current_string = Some(partial);
                            }
                        }
                        SstStringStage::Rich | SstStringStage::Ext => {
                            if !self.skip_aux(data, &mut offset, &mut partial) {
                                self.current_string = Some(partial);
                                return Ok(());
                            }
                        }
                    }
                }
            }
        }
    }

    /// 读取并解析字符串头部。头部完整返回 true，否则返回 false（等待更多数据）。
    fn read_header(
        &self,
        data: &[u8],
        offset: &mut usize,
        partial: &mut PartialString,
    ) -> bool {
        // 先补齐 cch + grbit（3 字节）
        if partial.buffer.len() < 3 {
            let take = (3 - partial.buffer.len()).min(data.len() - *offset);
            partial.buffer.extend_from_slice(&data[*offset..*offset + take]);
            *offset += take;
        }
        if partial.buffer.len() < 3 {
            return false;
        }

        let flag = partial.buffer[2];
        let has_rich = (flag & 0x08) != 0;
        let has_ext = (flag & 0x04) != 0;
        let need = 3 + if has_rich { 2 } else { 0 } + if has_ext { 4 } else { 0 };

        if partial.buffer.len() < need {
            let take = (need - partial.buffer.len()).min(data.len() - *offset);
            partial.buffer.extend_from_slice(&data[*offset..*offset + take]);
            *offset += take;
        }
        if partial.buffer.len() < need {
            return false;
        }

        partial.char_count = u16::from_le_bytes([partial.buffer[0], partial.buffer[1]]) as usize;
        partial.is_utf16 = (flag & 0x01) != 0;
        partial.has_rich = has_rich;
        partial.has_ext = has_ext;

        let mut pos = 3;
        if has_rich {
            partial.rich_runs =
                u16::from_le_bytes([partial.buffer[pos], partial.buffer[pos + 1]]);
            pos += 2;
        }
        if has_ext {
            partial.ext_size = u32::from_le_bytes([
                partial.buffer[pos],
                partial.buffer[pos + 1],
                partial.buffer[pos + 2],
                partial.buffer[pos + 3],
            ]);
        }

        partial.stage = SstStringStage::Chars;
        partial.bytes_read = 0;
        partial.buffer.clear();
        true
    }

    /// 读取字符数据。字符串字符完整读取并解码入表后返回 true，
    /// 跨块时返回 false 并保存进度（此时不修改 `current_string`）。
    fn read_chars(
        &mut self,
        data: &[u8],
        offset: &mut usize,
        partial: &mut PartialString,
        sst: &mut SharedStringTable,
    ) -> bool {
        let total = if partial.is_utf16 {
            partial.char_count * 2
        } else {
            partial.char_count
        };

        // 字节数足够则直接完成
        let available = data.len() - *offset;
        if partial.bytes_read + available < total {
            partial.buffer.extend_from_slice(&data[*offset..]);
            partial.bytes_read += available;
            *offset = data.len();
            return false;
        }

        let take = total.saturating_sub(partial.bytes_read);
        if take > 0 {
            partial.buffer.extend_from_slice(&data[*offset..*offset + take]);
            *offset += take;
        }
        partial.bytes_read = total;

        // 字符串完成，解码并加入 SST
        let s = Self::decode_string_parts(partial);
        sst.push_string(s);
        self.strings_parsed += 1;

        // 进入 Rich/Ext 跳过阶段（若有）
        partial.stage = if partial.has_rich && partial.rich_runs > 0 {
            SstStringStage::Rich
        } else if partial.has_ext && partial.ext_size > 0 {
            SstStringStage::Ext
        } else {
            SstStringStage::Header // 字符串完全结束
        };
        partial.bytes_read = 0;
        partial.buffer.clear();
        partial.segments.clear();
        true
    }

    /// 跳过 Rich Text runs 与扩展数据。全部跳过完成返回 true。
    fn skip_aux(
        &self,
        data: &[u8],
        offset: &mut usize,
        partial: &mut PartialString,
    ) -> bool {
        loop {
            let need_total = match partial.stage {
                SstStringStage::Rich => partial.rich_runs as usize * 4,
                SstStringStage::Ext => partial.ext_size as usize,
                _ => 0,
            };
            let remaining = need_total.saturating_sub(partial.bytes_read);
            let available = data.len() - *offset;

            if available < remaining {
                partial.bytes_read += available;
                *offset = data.len();
                return false;
            }

            *offset += remaining;
            partial.bytes_read = need_total;

            match partial.stage {
                SstStringStage::Rich if partial.has_ext => {
                    partial.stage = SstStringStage::Ext;
                    partial.bytes_read = 0;
                }
                _ => {
                    partial.stage = SstStringStage::Header;
                    return true;
                }
            }
        }
    }

    /// 完成解析，处理可能未完成的字符串
    pub fn finish(mut self, sst: &mut SharedStringTable) -> Result<(), XlsError> {
        if let Some(partial) = self.current_string.take() {
            if partial.stage == SstStringStage::Chars {
                // 尝试使用已有的数据解码（可能是截断的字符串）
                let s = Self::decode_string_parts(&partial);
                sst.push_string(s);
                self.strings_parsed += 1;
            }
        }
        Ok(())
    }

    /// 解码一个字符串（可能包含多个不同压缩模式的段）
    fn decode_string_parts(partial: &PartialString) -> String {
        let mut out = String::new();
        for (is_utf16, bytes) in partial
            .segments
            .iter()
            .chain(std::iter::once(&(partial.is_utf16, partial.buffer.clone())))
        {
            out.push_str(&Self::decode_string(bytes, *is_utf16));
        }
        out
    }

    fn decode_string(data: &[u8], is_utf16: bool) -> String {
        if is_utf16 {
            let u16_vec: Vec<u16> = data
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect();
            String::from_utf16(&u16_vec).unwrap_or_default()
        } else {
            String::from_utf8_lossy(data).to_string()
        }
    }
}

/// 解析状态
///
/// 在解析过程中维护的上下文状态
#[derive(Debug)]
pub struct ParseState {
    /// 当前正在解析的工作表
    pub current_sheet: Option<XlsSheet>,
    /// 已完成的全部工作表
    pub sheets: Vec<XlsSheet>,
    /// 共享字符串表
    pub sst: SharedStringTable,
    /// 工作表名称列表（从 BOUNDSHEET 收集）
    pub sheet_names: Vec<String>,
    /// SST 解析状态机
    pub sst_parser: Option<SSTParserState>,
    /// 标记解析是否完成
    pub is_complete: bool,
    /// 上一个字符串结果公式的行列位置（等待 STRING 记录）
    pub last_formula_string_cell: Option<(u16, u16)>,
}

impl ParseState {
    /// 创建新的解析状态
    pub fn new() -> Self {
        ParseState {
            current_sheet: None,
            sheets: Vec::new(),
            sst: SharedStringTable::new(),
            sheet_names: Vec::new(),
            sst_parser: None,
            is_complete: false,
            last_formula_string_cell: None,
        }
    }

    /// 获取当前工作表的可变引用
    ///
    /// 如果当前没有工作表，返回错误
    pub fn current_sheet_mut(&mut self) -> Result<&mut XlsSheet, XlsError> {
        self.current_sheet
            .as_mut()
            .ok_or_else(|| XlsError::InvalidFormat("No current sheet available".to_string()))
    }
}

impl Default for ParseState {
    fn default() -> Self {
        Self::new()
    }
}

/// RK 值解码
///
/// 将 RK 编码的 32 位整数解码为 f64
///
/// # RK 编码格式
/// - bit 0: 是否为除以 100 的值
/// - bit 1: 是否为整数（1=整数，0=浮点数）
/// - bits 2-31: 实际数值（整数时右移 2 位，浮点数时作为 f64 尾数）
pub fn decode_rk_value(rk: i32) -> f64 {
    let is_div_100 = (rk & 0x01) != 0;
    let is_int = (rk & 0x02) != 0;

    let val = if is_int {
        ((rk >> 2) as i32) as f64
    } else {
        // 浮点数：低 30 位左移 32 位作为 f64 尾数
        let bits = ((rk as u32) & 0xFFFF_FFFC) as u64;
        f64::from_bits(bits << 32)
    };

    if is_div_100 {
        val / 100.0
    } else {
        val
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_decode_rk_integer() {
        // 编码整数 42: (42 << 2) | 0x02 = 170
        let encoded = (42_i32 << 2) | 0x02;
        assert_eq!(decode_rk_value(encoded), 42.0);
    }

    #[test]
    fn test_decode_rk_integer_div100() {
        // 编码整数 42，除以 100: (42 << 2) | 0x03 = 171
        let encoded = (42_i32 << 2) | 0x03;
        assert_eq!(decode_rk_value(encoded), 0.42);
    }

    #[test]
    fn test_parse_state_new() {
        let state = ParseState::new();
        assert!(state.current_sheet.is_none());
        assert!(state.sheets.is_empty());
        assert!(state.sheet_names.is_empty());
        assert_eq!(state.sst.string_count(), 0);
    }
}
