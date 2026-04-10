use super::BiffRecord;

/// EOF (End of File) 记录
///
/// ## 作用
///
/// 标识BIFF段落的结束。每个BIFF段（如工作簿globals或worksheet）结束时需要此记录。
/// EOF记录没有数据部分，data()返回空向量。
#[derive(Debug, Default)]
pub struct EofRecord;

impl EofRecord {
    pub fn new() -> Self {
        EofRecord
    }
}

impl BiffRecord for EofRecord {
    fn id(&self) -> u16 {
        0x000A
    }

    fn data(&self) -> Vec<u8> {
        Vec::new()
    }
}

// ============================================================================
// ParsableRecord implementation for reading
// ============================================================================

use crate::xls::records::{ParsableRecord, ParseState};
use crate::xls::XlsError;

impl ParsableRecord for EofRecord {
    const RECORD_ID: u16 = 0x000A;

    fn parse(_data: &[u8]) -> Result<Self, XlsError> {
        Ok(EofRecord::new())
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        // 完成 SST
        if let Some(parser) = state.sst_parser.take() {
            if let Err(e) = parser.finish(&mut state.sst) {
                eprintln!("Warning: Failed to finish SST on EOF: {}", e);
            }
        }

        if let Some(sheet) = state.current_sheet.take() {
            state.sheets.push(sheet);
        }

        // 检查是否完成
        if state.sheets.len() >= state.sheet_names.len() && !state.sheet_names.is_empty() {
            state.is_complete = true;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_eof_record_id() {
        let record = EofRecord::new();
        assert_eq!(record.id(), 0x000A);
    }

    #[test]
    fn test_eof_record_data_size() {
        let record = EofRecord::new();
        assert_eq!(record.data().len(), 0);
    }
}
