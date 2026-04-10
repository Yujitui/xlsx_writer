use crate::xls::records::BiffRecord;
use crate::xls::XlsCell;

/// DimensionsRecord 记录
///
/// 作用：定义工作表的数据范围
///
/// DimensionsRecord是Excel BIFF格式中的维度记录（ID: 0x0200），用于定义工作表中
/// 使用单元格的范围。这个信息帮助Excel快速定位有效数据区域。
///
/// **注意**：last_used_row和last_used_col存储的是"最后一个使用的位置+1"
///
/// ## 参数说明
///
/// - `first_used_row`: 第一个使用的行索引（0-based）
/// - `last_used_row`: 最后一个使用的行索引 + 1
/// - `first_used_col`: 第一个使用的列索引（0-based）
/// - `last_used_col`: 最后一个使用的列索引 + 1
#[derive(Debug)]
pub struct DimensionsRecord {
    first_used_row: u32,
    last_used_row: u32,
    first_used_col: u16,
    last_used_col: u16,
}

impl DimensionsRecord {
    pub fn new(
        first_used_row: u32,
        last_used_row: u32,
        first_used_col: u16,
        last_used_col: u16,
    ) -> Self {
        if first_used_row > last_used_row || first_used_col > last_used_col {
            DimensionsRecord {
                first_used_row: 0,
                last_used_row: 0,
                first_used_col: 0,
                last_used_col: 0,
            }
        } else {
            DimensionsRecord {
                first_used_row,
                last_used_row: last_used_row + 1,
                first_used_col,
                last_used_col: last_used_col + 1,
            }
        }
    }
}

impl Default for DimensionsRecord {
    fn default() -> Self {
        DimensionsRecord {
            first_used_row: 0,
            last_used_row: 0,
            first_used_col: 0,
            last_used_col: 0,
        }
    }
}

impl DimensionsRecord {
    pub fn from(data: &Vec<Vec<Option<XlsCell>>>) -> Self {
        if data.is_empty() {
            return DimensionsRecord::default();
        }

        let first_used_row = 0u32;
        let last_used_row = (data.len() - 1) as u32;

        let first_used_col = data
            .iter()
            .enumerate()
            .flat_map(|(row_idx, row)| {
                row.iter()
                    .enumerate()
                    .map(move |(col_idx, c)| (row_idx, col_idx, c))
            })
            .find(|(_, _, c)| c.is_some())
            .map(|(_, col_idx, _)| col_idx as u16)
            .unwrap_or(0);

        let last_used_col = data
            .iter()
            .enumerate()
            .flat_map(|(row_idx, row)| {
                row.iter()
                    .enumerate()
                    .map(move |(col_idx, c)| (row_idx, col_idx, c))
            })
            .rfind(|(_, _, c)| c.is_some())
            .map(|(_, col_idx, _)| col_idx as u16)
            .unwrap_or(0);

        DimensionsRecord::new(first_used_row, last_used_row, first_used_col, last_used_col)
    }
}

impl BiffRecord for DimensionsRecord {
    fn id(&self) -> u16 {
        0x0200
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(14);
        buf.extend_from_slice(&self.first_used_row.to_le_bytes());
        buf.extend_from_slice(&self.last_used_row.to_le_bytes());
        buf.extend_from_slice(&self.first_used_col.to_le_bytes());
        buf.extend_from_slice(&self.last_used_col.to_le_bytes());
        buf.extend_from_slice(&0x0000u16.to_le_bytes());
        buf
    }
}

// ============================================================================
// ParsableRecord implementation for reading
// ============================================================================

use crate::xls::records::{ParsableRecord, ParseState};
use crate::xls::XlsError;

impl ParsableRecord for DimensionsRecord {
    const RECORD_ID: u16 = 0x0200;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 14 {
            return Err(XlsError::InvalidFormat(format!(
                "DimensionsRecord data too short: {} bytes",
                data.len()
            )));
        }

        let first_used_row = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
        let last_used_row = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
        let first_used_col = u16::from_le_bytes([data[8], data[9]]);
        let last_used_col = u16::from_le_bytes([data[10], data[11]]);

        Ok(DimensionsRecord {
            first_used_row,
            last_used_row,
            first_used_col,
            last_used_col,
        })
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        let sheet = state.current_sheet_mut()?;

        // 预分配行空间
        // 注意：last_used_row 存储的是 "最后一个使用的行 + 1"
        let num_rows = if self.last_used_row > self.first_used_row {
            (self.last_used_row - self.first_used_row) as usize
        } else {
            0
        };

        // 预分配列空间
        let num_cols = if self.last_used_col > self.first_used_col {
            (self.last_used_col - self.first_used_col) as usize
        } else {
            0
        };

        // 预分配行
        for _ in 0..num_rows {
            sheet.rows.push(vec![None; num_cols]);
        }

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_dimensions_record_id() {
        let record = DimensionsRecord::default();
        assert_eq!(record.id(), 0x0200);
    }

    #[test]
    fn test_dimensions_record_normal_case() {
        let record = DimensionsRecord::new(0, 10, 0, 5);
        let data = record.data();

        assert_eq!(&data[0..4], &0u32.to_le_bytes());
        assert_eq!(&data[4..8], &11u32.to_le_bytes());
        assert_eq!(&data[8..10], &0u16.to_le_bytes());
        assert_eq!(&data[10..12], &6u16.to_le_bytes());
    }

    #[test]
    fn test_dimensions_record_empty_worksheet() {
        let record = DimensionsRecord::new(65535, 0, 255, 0);
        let data = record.data();

        assert_eq!(&data[0..4], &0u32.to_le_bytes());
        assert_eq!(&data[4..8], &0u32.to_le_bytes());
        assert_eq!(&data[8..10], &0u16.to_le_bytes());
        assert_eq!(&data[10..12], &0u16.to_le_bytes());
    }

    #[test]
    fn test_dimensions_record_data_size() {
        let record = DimensionsRecord::default();
        assert_eq!(record.data().len(), 14);
    }
}
