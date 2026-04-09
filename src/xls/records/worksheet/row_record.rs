use crate::xls::records::BiffRecord;
use crate::xls::XlsCell;

/// RowRecord 记录
///
/// 作用：存储行的属性信息
///
/// RowRecord是Excel BIFF格式中的行记录（ID: 0x0208），用于定义工作表中
/// 某一行的属性信息，如行高、是否折叠、格式索引等。
///
/// **注意**：当行没有数据时，此记录的serialize方法返回空数据，不写入文件。
///
/// ## 参数说明
///
/// - `index`: 行索引（0-based）
/// - `first_col`: 第一个有数据的列索引
/// - `last_col`: 最后一个有数据的列索引 + 1
/// - `height_options`: 行高选项
/// - `options`: 行选项标志
#[derive(Debug)]
pub struct RowRecord {
    index: u16,
    first_col: u16,
    last_col: u16,
    height_options: u16,
    options: u32,
}

impl RowRecord {
    pub fn from_row_data(row_index: usize, row_data: &Vec<Option<XlsCell>>) -> Self {
        let first_col = row_data.iter().position(|c| c.is_some()).unwrap_or(0) as u16;

        let last_col = row_data
            .iter()
            .rposition(|c| c.is_some())
            .map(|p| p + 1)
            .unwrap_or(0) as u16;

        let height_options = 0x00FF;
        let options = 0x000F0100;

        RowRecord {
            index: row_index as u16,
            first_col,
            last_col,
            height_options,
            options,
        }
    }
}

impl BiffRecord for RowRecord {
    fn id(&self) -> u16 {
        0x0208
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(16);
        buf.extend_from_slice(&self.index.to_le_bytes());
        buf.extend_from_slice(&self.first_col.to_le_bytes());
        buf.extend_from_slice(&self.last_col.to_le_bytes());
        buf.extend_from_slice(&self.height_options.to_le_bytes());
        buf.extend_from_slice(&0x0000u16.to_le_bytes());
        buf.extend_from_slice(&0x0000u16.to_le_bytes());
        buf.extend_from_slice(&self.options.to_le_bytes());
        buf
    }

    fn serialize(&self) -> Vec<u8> {
        if self.first_col == 0 && self.last_col == 0 {
            vec![]
        } else {
            BiffRecord::serialize(self)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xls::XlsCell;

    #[test]
    fn test_row_record_id() {
        let row_data = vec![Some(XlsCell::Text("A".to_string()))];
        let record = RowRecord::from_row_data(0, &row_data);
        assert_eq!(record.id(), 0x0208);
    }

    #[test]
    fn test_row_record_data_size() {
        let row_data = vec![Some(XlsCell::Text("A".to_string()))];
        let record = RowRecord::from_row_data(0, &row_data);
        assert_eq!(record.data().len(), 16);
    }

    #[test]
    fn test_row_record_from_row_data_with_data() {
        let row_data = vec![
            Some(XlsCell::Text("A".to_string())),
            Some(XlsCell::Text("B".to_string())),
            None,
            Some(XlsCell::Text("D".to_string())),
        ];
        let record = RowRecord::from_row_data(0, &row_data);

        assert_eq!(record.index, 0);
        assert_eq!(record.first_col, 0);
        assert_eq!(record.last_col, 4);
    }

    #[test]
    fn test_row_record_from_row_data_all_empty() {
        let row_data: Vec<Option<XlsCell>> = vec![None, None, None];
        let record = RowRecord::from_row_data(5, &row_data);

        let serialized = record.serialize();
        assert!(serialized.is_empty());
    }

    #[test]
    fn test_row_record_from_row_data_with_row_index() {
        let row_data = vec![Some(XlsCell::Text("A".to_string()))];
        let record = RowRecord::from_row_data(10, &row_data);

        assert_eq!(record.index, 10);
    }
}
