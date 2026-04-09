use crate::xls::records::BiffRecord;

/// CellRange 结构体
///
/// ## 作用
///
/// 表示一个合并的单元格区域。
///
/// ## 参数说明
///
/// - `first_row`: 起始行索引（0-based）
/// - `last_row`: 结束行索引（0-based）
/// - `first_col`: 起始列索引（0-based）
/// - `last_col`: 结束列索引（0-based）
#[derive(Debug, Clone)]
pub struct CellRange {
    pub first_row: u16,
    pub last_row: u16,
    pub first_col: u16,
    pub last_col: u16,
}

/// MergedCellsRecord 记录
///
/// 作用：存储工作表中的合并单元格信息
///
/// MergedCellsRecord是Excel BIFF格式中的合并单元格记录（ID: 0x00E5），用于定义
/// 工作表中已合并的单元格区域。每个合并区域由左上角和右下角的单元格坐标定义。
///
/// **注意**：当没有合并单元格时，此记录的serialize方法返回空数据，不写入文件。
///
/// ## 参数说明
///
/// - `ranges`: 合并单元格区域向量
#[derive(Debug)]
pub struct MergedCellsRecord {
    ranges: Vec<CellRange>,
}

impl MergedCellsRecord {
    pub fn new(ranges: Vec<CellRange>) -> Self {
        MergedCellsRecord { ranges }
    }

    pub fn from_tuples(ranges: Vec<(u16, u16, u16, u16)>) -> Self {
        let cell_ranges = ranges
            .into_iter()
            .map(|(first_row, last_row, first_col, last_col)| CellRange {
                first_row,
                last_row,
                first_col,
                last_col,
            })
            .collect();
        MergedCellsRecord {
            ranges: cell_ranges,
        }
    }
}

impl Default for MergedCellsRecord {
    fn default() -> Self {
        MergedCellsRecord { ranges: Vec::new() }
    }
}

impl BiffRecord for MergedCellsRecord {
    fn id(&self) -> u16 {
        0x00E5
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::new();
        buf.extend_from_slice(&(self.ranges.len() as u16).to_le_bytes());
        for range in &self.ranges {
            buf.extend_from_slice(&range.first_row.to_le_bytes());
            buf.extend_from_slice(&range.last_row.to_le_bytes());
            buf.extend_from_slice(&range.first_col.to_le_bytes());
            buf.extend_from_slice(&range.last_col.to_le_bytes());
        }
        buf
    }

    fn serialize(&self) -> Vec<u8> {
        if self.ranges.is_empty() {
            return Vec::new();
        }
        BiffRecord::serialize(self)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_merged_cells_record_id() {
        let record = MergedCellsRecord::default();
        assert_eq!(record.id(), 0x00E5);
    }

    #[test]
    fn test_merged_cells_record_empty() {
        let record = MergedCellsRecord::default();
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_merged_cells_record_empty_serialize() {
        let record = MergedCellsRecord::default();
        let serialized = record.serialize();
        assert!(serialized.is_empty());
    }

    #[test]
    fn test_merged_cells_record_single_range() {
        let ranges = vec![CellRange {
            first_row: 0,
            last_row: 1,
            first_col: 0,
            last_col: 1,
        }];
        let record = MergedCellsRecord::new(ranges);
        let data = record.data();

        assert_eq!(&data[0..2], &1u16.to_le_bytes());
        assert_eq!(&data[2..4], &0u16.to_le_bytes());
        assert_eq!(&data[4..6], &1u16.to_le_bytes());
        assert_eq!(&data[6..8], &0u16.to_le_bytes());
        assert_eq!(&data[8..10], &1u16.to_le_bytes());
    }

    #[test]
    fn test_merged_cells_record_from_tuples() {
        let ranges = vec![(0, 1, 0, 1), (3, 4, 2, 3)];
        let record = MergedCellsRecord::from_tuples(ranges);
        let data = record.data();

        assert_eq!(data.len(), 18);
    }

    #[test]
    fn test_merged_cells_record_multiple_ranges() {
        let ranges = vec![
            CellRange {
                first_row: 0,
                last_row: 1,
                first_col: 0,
                last_col: 1,
            },
            CellRange {
                first_row: 3,
                last_row: 4,
                first_col: 2,
                last_col: 3,
            },
        ];
        let record = MergedCellsRecord::new(ranges);
        let data = record.data();

        assert_eq!(&data[0..2], &2u16.to_le_bytes());
        assert_eq!(data.len(), 18);
    }
}
