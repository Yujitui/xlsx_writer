//! 可解析记录 trait
//!
//! 为 BIFF 记录提供统一的反序列化接口

use crate::xls::records::workbook::sst_record::SharedStringTable;
use crate::xls::{XlsError, XlsSheet};

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

/// 跨边界部分字符串状态
#[derive(Debug)]
pub struct PartialString {
    pub char_count: usize,
    pub is_utf16: bool,
    pub has_rich: bool,
    pub has_ext: bool,
    pub rich_runs: u16,
    pub ext_size: u32,
    pub bytes_read: usize,
    pub buffer: Vec<u8>,
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
    pub fn parse_chunk(
        &mut self,
        data: &[u8],
        sst: &mut SharedStringTable,
    ) -> Result<(), XlsError> {
        let mut offset = 0;

        // 如果有未完成的字符串，先尝试完成
        if let Some(ref mut partial) = self.current_string {
            let needed = if partial.is_utf16 {
                partial.char_count * 2 - partial.bytes_read
            } else {
                partial.char_count - partial.bytes_read
            };

            let available = data.len().min(needed);
            partial.buffer.extend_from_slice(&data[..available]);
            partial.bytes_read += available;
            offset += available;

            let total_needed = if partial.is_utf16 {
                partial.char_count * 2
            } else {
                partial.char_count
            };

            if partial.bytes_read >= total_needed {
                // 字符串完成，添加到 SST
                let s = Self::decode_string(&partial.buffer, partial.is_utf16);
                sst.push_string(s);
                self.strings_parsed += 1;
                self.current_string = None;
            } else {
                // 还需要更多数据
                return Ok(());
            }
        }

        // 解析完整的字符串
        while self.strings_parsed < self.unique_count && offset < data.len() {
            // 检查是否有足够数据读取头部（3 字节）
            if offset + 3 > data.len() {
                return Ok(());
            }

            let char_count = u16::from_le_bytes([data[offset], data[offset + 1]]) as usize;
            let flag = data[offset + 2];
            offset += 3;

            let is_utf16 = (flag & 0x01) != 0;
            let has_rich = (flag & 0x08) != 0;
            let has_ext = (flag & 0x04) != 0;

            // 读取 Rich/Ext 信息
            let mut rich_runs = 0u16;
            let mut ext_size = 0u32;

            if has_rich {
                if offset + 2 > data.len() {
                    // 保存状态到下一次
                    self.current_string = Some(PartialString {
                        char_count,
                        is_utf16,
                        has_rich,
                        has_ext,
                        rich_runs: 0,
                        ext_size: 0,
                        bytes_read: 0,
                        buffer: Vec::new(),
                    });
                    return Ok(());
                }
                rich_runs = u16::from_le_bytes([data[offset], data[offset + 1]]);
                offset += 2;
            }

            if has_ext {
                if offset + 4 > data.len() {
                    self.current_string = Some(PartialString {
                        char_count,
                        is_utf16,
                        has_rich,
                        has_ext,
                        rich_runs,
                        ext_size: 0,
                        bytes_read: 0,
                        buffer: Vec::new(),
                    });
                    return Ok(());
                }
                ext_size = u32::from_le_bytes([
                    data[offset],
                    data[offset + 1],
                    data[offset + 2],
                    data[offset + 3],
                ]);
                offset += 4;
            }

            // 尝试读取字符串
            let string_bytes = if is_utf16 { char_count * 2 } else { char_count };

            if offset + string_bytes > data.len() {
                // 字符串跨边界，创建 PartialString
                let available = data.len() - offset;
                self.current_string = Some(PartialString {
                    char_count,
                    is_utf16,
                    has_rich,
                    has_ext,
                    rich_runs,
                    ext_size,
                    bytes_read: available,
                    buffer: data[offset..].to_vec(),
                });
                return Ok(());
            }

            // 完整字符串，直接解析
            let s = Self::decode_string(&data[offset..offset + string_bytes], is_utf16);
            sst.push_string(s);
            self.strings_parsed += 1;

            offset += string_bytes;
            // 跳过 Rich Text 和 Extension 数据
            offset += (rich_runs as usize * 4) + ext_size as usize;
        }

        Ok(())
    }

    /// 完成解析，处理可能未完成的字符串
    pub fn finish(mut self, sst: &mut SharedStringTable) -> Result<(), XlsError> {
        if let Some(partial) = self.current_string.take() {
            // 尝试使用已有的数据解码
            let s = Self::decode_string(&partial.buffer, partial.is_utf16);
            sst.push_string(s);
            self.strings_parsed += 1;
        }
        Ok(())
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
