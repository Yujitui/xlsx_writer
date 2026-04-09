use crate::xls::records::workbook::sst_record::SharedStringTable;
use crate::xls::records::BiffRecord;
use crate::xls::XlsCell;

/// LabelSSTRecord 记录
///
/// 作用：存储字符串单元格数据（引用共享字符串表）
///
/// LabelSSTRecord是Excel BIFF格式中的字符串标签记录（ID: 0x00FD），用于存储包含文本的单元格。
/// 实际字符串数据存储在共享字符串表（SST）中，这里只存储对字符串的引用索引。
///
/// ## 参数说明
///
/// - `row`: 行索引（0-based）
/// - `col`: 列索引（0-based）
/// - `xf_idx`: 格式索引（指向XFRecord）
/// - `sst_idx`: 共享字符串表索引（指向SST中的字符串）
#[derive(Debug)]
pub struct LabelSSTRecord {
    row: u16,
    col: u16,
    xf_idx: u16,
    sst_idx: u32,
}

impl LabelSSTRecord {
    pub fn new(row: u16, col: u16, xf_idx: u16, sst_idx: u32) -> Self {
        LabelSSTRecord {
            row,
            col,
            xf_idx,
            sst_idx,
        }
    }
}

impl BiffRecord for LabelSSTRecord {
    fn id(&self) -> u16 {
        0x00FD
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(10);
        buf.extend_from_slice(&self.row.to_le_bytes());
        buf.extend_from_slice(&self.col.to_le_bytes());
        buf.extend_from_slice(&self.xf_idx.to_le_bytes());
        buf.extend_from_slice(&self.sst_idx.to_le_bytes());
        buf
    }
}

/// BlankRecord 记录
///
/// 作用：存储空白单元格（无数据但有格式）
///
/// BlankRecord是Excel BIFF格式中的空白单元格记录（ID: 0x0201），用于表示
/// 没有数据但有格式样式的单元格。常用于设置列宽/行高后填补空白。
///
/// ## 参数说明
///
/// - `row`: 行索引（0-based）
/// - `col`: 列索引（0-based）
/// - `xf_idx`: 格式索引（指向XFRecord）
#[derive(Debug)]
pub struct BlankRecord {
    row: u16,
    col: u16,
    xf_idx: u16,
}

impl BlankRecord {
    pub fn new(row: u16, col: u16, xf_idx: u16) -> Self {
        BlankRecord { row, col, xf_idx }
    }
}

impl BiffRecord for BlankRecord {
    fn id(&self) -> u16 {
        0x0201
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(6);
        buf.extend_from_slice(&self.row.to_le_bytes());
        buf.extend_from_slice(&self.col.to_le_bytes());
        buf.extend_from_slice(&self.xf_idx.to_le_bytes());
        buf
    }
}

/// MulBlankRecord 记录
///
/// 作用：存储连续多个空白单元格（优化记录）
///
/// MulBlankRecord是Excel BIFF格式中的多空白单元格记录（ID: 0x00BE），用于优化存储
/// 连续多个具有相同格式的空白单元格。相比逐个创建BlankRecord，MulBlankRecord可以
/// 减少记录数量，减小文件体积。
///
/// ## 参数说明
///
/// - `row`: 行索引（0-based）
/// - `first_col`: 起始列索引（0-based）
/// - `last_col`: 结束列索引（0-based）
/// - `xf_idx`: 格式索引（指向XFRecord，所有空白单元格使用相同格式）
#[derive(Debug)]
pub struct MulBlankRecord {
    row: u16,
    first_col: u16,
    last_col: u16,
    xf_idx: u16,
}

impl MulBlankRecord {
    pub fn new(row: u16, first_col: u16, last_col: u16, xf_idx: u16) -> Self {
        MulBlankRecord {
            row,
            first_col,
            last_col,
            xf_idx,
        }
    }
}

impl BiffRecord for MulBlankRecord {
    fn id(&self) -> u16 {
        0x00BE
    }

    fn data(&self) -> Vec<u8> {
        let count = self.last_col - self.first_col + 1;
        let mut buf = Vec::with_capacity(4 + count as usize * 2);
        buf.extend_from_slice(&self.row.to_le_bytes());
        buf.extend_from_slice(&self.first_col.to_le_bytes());
        for _ in 0..count {
            buf.extend_from_slice(&self.xf_idx.to_le_bytes());
        }
        buf.extend_from_slice(&self.last_col.to_le_bytes());
        buf
    }
}

/// NumberRecord 记录
///
/// 作用：存储浮点数单元格数据
///
/// NumberRecord是Excel BIFF格式中的数字记录（ID: 0x0203），用于存储浮点数值的单元格。
/// 当数值无法用RKRecord的压缩整数格式表示时使用此记录。
///
/// ## 参数说明
///
/// - `row`: 行索引（0-based）
/// - `col`: 列索引（0-based）
/// - `xf_idx`: 格式索引（指向XFRecord）
/// - `number`: 浮点数值
#[derive(Debug)]
pub struct NumberRecord {
    row: u16,
    col: u16,
    xf_idx: u16,
    number: f64,
}

impl NumberRecord {
    pub fn new(row: u16, col: u16, xf_idx: u16, number: f64) -> Self {
        NumberRecord {
            row,
            col,
            xf_idx,
            number,
        }
    }
}

impl BiffRecord for NumberRecord {
    fn id(&self) -> u16 {
        0x0203
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(14);
        buf.extend_from_slice(&self.row.to_le_bytes());
        buf.extend_from_slice(&self.col.to_le_bytes());
        buf.extend_from_slice(&self.xf_idx.to_le_bytes());
        buf.extend_from_slice(&self.number.to_le_bytes());
        buf
    }
}

/// RKRecord 记录
///
/// 作用：存储RK编码的整数/小数单元格数据（压缩格式）
///
/// RKRecord是Excel BIFF格式中的RK数字记录（ID: 0x027E），用于存储可以用
/// 压缩格式表示的整数或乘以100的整数。这种格式可以减少文件大小。
///
/// ## 参数说明
///
/// - `row`: 行索引（0-based）
/// - `col`: 列索引（0-based）
/// - `xf_idx`: 格式索引（指向XFRecord）
/// - `rk_encoded`: RK编码的整数（包含标志位 + 数值）
///   - bit 0: 值类型（0=整数, 1=乘以100的小数）
///   - bit 1-31: 有符号整数（需要右移2位获取实际值）
#[derive(Debug)]
pub struct RKRecord {
    row: u16,
    col: u16,
    xf_idx: u16,
    rk_encoded: i32,
}

impl RKRecord {
    pub fn new(row: u16, col: u16, xf_idx: u16, rk_encoded: i32) -> Self {
        RKRecord {
            row,
            col,
            xf_idx,
            rk_encoded,
        }
    }
}

impl BiffRecord for RKRecord {
    fn id(&self) -> u16 {
        0x027E
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(10);
        buf.extend_from_slice(&self.row.to_le_bytes());
        buf.extend_from_slice(&self.col.to_le_bytes());
        buf.extend_from_slice(&self.xf_idx.to_le_bytes());
        buf.extend_from_slice(&self.rk_encoded.to_le_bytes());
        buf
    }
}

/// MulRkRecord 记录
///
/// 作用：存储连续多个RK编码的数字单元格（优化记录）
///
/// MulRkRecord是Excel BIFF格式中的多RK数字记录（ID: 0x00BD），用于优化存储
/// 连续多个可以用RK编码表示的数字。相比逐个创建RKRecord，MulRkRecord可以减少记录数量。
///
/// ## 参数说明
///
/// - `row`: 行索引（0-based）
/// - `first_col`: 起始列索引（0-based）
/// - `last_col`: 结束列索引（0-based）
/// - `rk_values`: RK值向量，每个元素为(列索引, RK编码值)的元组
#[derive(Debug)]
pub struct MulRkRecord {
    row: u16,
    first_col: u16,
    last_col: u16,
    rk_values: Vec<(u16, i32)>,
}

impl MulRkRecord {
    pub fn new(row: u16, first_col: u16, last_col: u16, rk_values: Vec<(u16, i32)>) -> Self {
        MulRkRecord {
            row,
            first_col,
            last_col,
            rk_values,
        }
    }
}

impl BiffRecord for MulRkRecord {
    fn id(&self) -> u16 {
        0x00BD
    }

    fn data(&self) -> Vec<u8> {
        let count = self.last_col - self.first_col + 1;
        let mut buf = Vec::with_capacity(4 + count as usize * 6);
        buf.extend_from_slice(&self.row.to_le_bytes());
        buf.extend_from_slice(&self.first_col.to_le_bytes());
        for (xf_idx, rk_encoded) in &self.rk_values {
            buf.extend_from_slice(&xf_idx.to_le_bytes());
            buf.extend_from_slice(&rk_encoded.to_le_bytes());
        }
        buf.extend_from_slice(&self.last_col.to_le_bytes());
        buf
    }
}

/// BoolErrRecord 记录
///
/// 作用：存储布尔值或错误代码单元格数据
///
/// BoolErrRecord是Excel BIFF格式中的布尔/错误记录（ID: 0x0205），用于存储
/// 布尔值（TRUE/FALSE）或错误值（#N/A, #REF!, #DIV/0!等）。
///
/// ## 参数说明
///
/// - `row`: 行索引（0-based）
/// - `col`: 列索引（0-based）
/// - `xf_idx`: 格式索引（指向XFRecord）
/// - `value`: 布尔值（0=FALSE, 1=TRUE）或错误代码
/// - `is_error`: 标志位（0=布尔值, 1=错误值）
#[derive(Debug)]
pub struct BoolErrRecord {
    row: u16,
    col: u16,
    xf_idx: u16,
    value: u8,
    is_error: u8,
}

impl BoolErrRecord {
    pub fn new(row: u16, col: u16, xf_idx: u16, value: u8, is_error: u8) -> Self {
        BoolErrRecord {
            row,
            col,
            xf_idx,
            value,
            is_error,
        }
    }
}

impl BiffRecord for BoolErrRecord {
    fn id(&self) -> u16 {
        0x0205
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(8);
        buf.extend_from_slice(&self.row.to_le_bytes());
        buf.extend_from_slice(&self.col.to_le_bytes());
        buf.extend_from_slice(&self.xf_idx.to_le_bytes());
        buf.push(self.value);
        buf.push(self.is_error);
        buf
    }
}

fn encode_rk_value(num: f64) -> Option<i32> {
    if -536870912.0 <= num && num < 536870912.0 {
        let inum = num as i32;
        if (inum as f64) == num {
            return Some(2 | (inum << 2));
        }
    }

    let temp = num * 100.0;
    if -536870912.0 <= temp && temp < 536870912.0 {
        let itemp = temp.round() as i32;
        if (itemp as f64) / 100.0 == num {
            return Some(3 | (itemp << 2));
        }
    }

    None
}

pub fn row_data_to_cell_records(
    row_index: usize,
    row_data: &Vec<Option<XlsCell>>,
    xf_index: u16,
    sst: &mut SharedStringTable,
) -> Vec<u8> {
    let mut result = Vec::new();
    let row = row_index as u16;
    let n = row_data.len();
    let mut i = 0;

    while i < n {
        let col = i as u16;
        let cell = &row_data[i];

        match cell {
            Some(XlsCell::Text(s)) => {
                let sst_idx = sst.add(s.clone()) as u32;
                result.extend_from_slice(
                    &LabelSSTRecord::new(row, col, xf_index, sst_idx).serialize(),
                );
                i += 1;
            }
            Some(XlsCell::Number(num)) => {
                // 尝试找到连续的可 RK 编码数字
                let mut rk_values: Vec<(u16, i32)> = Vec::new();
                let mut j = i;

                while j < n {
                    if let Some(XlsCell::Number(n)) = &row_data[j] {
                        if let Some(rk_encoded) = encode_rk_value(*n) {
                            rk_values.push((j as u16, rk_encoded));
                            j += 1;
                        } else {
                            break;
                        }
                    } else {
                        break;
                    }
                }

                let count = j - i;

                if count == 1 {
                    // 单个数字
                    if let Some(rk_encoded) = encode_rk_value(*num) {
                        result.extend_from_slice(
                            &RKRecord::new(row, col, xf_index, rk_encoded).serialize(),
                        );
                    } else {
                        result.extend_from_slice(
                            &NumberRecord::new(row, col, xf_index, *num).serialize(),
                        );
                    }
                } else {
                    // 多个连续的可 RK 编码数字
                    let first_col = col;
                    let last_col = (j - 1) as u16;

                    let mut rk_list: Vec<(u16, i32)> = Vec::new();
                    for k in i..j {
                        if let Some(XlsCell::Number(n)) = &row_data[k] {
                            if let Some(rk) = encode_rk_value(*n) {
                                rk_list.push((xf_index, rk));
                            }
                        }
                    }

                    result.extend_from_slice(
                        &MulRkRecord::new(row, first_col, last_col, rk_list).serialize(),
                    );
                }
                i = j;
            }
            Some(XlsCell::Boolean(b)) => {
                result.extend_from_slice(
                    &BoolErrRecord::new(row, col, xf_index, if *b { 1 } else { 0 }, 0).serialize(),
                );
                i += 1;
            }
            None => {
                let mut j = i + 1;
                while j < n && row_data[j].is_none() {
                    j += 1;
                }
                let last_col = (j - 1) as u16;
                if last_col == col {
                    result.extend_from_slice(&BlankRecord::new(row, col, xf_index).serialize());
                } else {
                    result.extend_from_slice(
                        &MulBlankRecord::new(row, col, last_col, xf_index).serialize(),
                    );
                }
                i = j;
            }
        }
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_label_sst_record_id() {
        let record = LabelSSTRecord::new(0, 0, 0x0F, 0);
        assert_eq!(record.id(), 0x00FD);
    }

    #[test]
    fn test_label_sst_record_data_size() {
        let record = LabelSSTRecord::new(0, 0, 0x0F, 0);
        assert_eq!(record.data().len(), 10);
    }

    #[test]
    fn test_blank_record_id() {
        let record = BlankRecord::new(0, 0, 0x0F);
        assert_eq!(record.id(), 0x0201);
    }

    #[test]
    fn test_blank_record_data_size() {
        let record = BlankRecord::new(0, 0, 0x0F);
        assert_eq!(record.data().len(), 6);
    }

    #[test]
    fn test_mul_blank_record() {
        let record = MulBlankRecord::new(0, 0, 4, 0x0F);
        let data = record.data();
        assert_eq!(record.id(), 0x00BE);
        assert_eq!(data.len(), 16);
    }

    #[test]
    fn test_number_record_id() {
        let record = NumberRecord::new(0, 0, 0x0F, 1.0);
        assert_eq!(record.id(), 0x0203);
    }

    #[test]
    fn test_number_record_data_size() {
        let record = NumberRecord::new(0, 0, 0x0F, 1.0);
        assert_eq!(record.data().len(), 14);
    }

    #[test]
    fn test_rk_record() {
        let record = RKRecord::new(0, 0, 0x0F, 4);
        assert_eq!(record.id(), 0x027E);
        assert_eq!(record.data().len(), 10);
    }

    #[test]
    fn test_encode_rk_value_integer() {
        assert!(encode_rk_value(42.0).is_some());
    }

    #[test]
    fn test_encode_rk_value_integer_times_100() {
        assert!(encode_rk_value(42.55).is_some());
    }

    #[test]
    fn test_encode_rk_value_float_not_rk() {
        assert!(encode_rk_value(std::f64::consts::PI).is_none());
    }

    #[test]
    fn test_bool_err_record() {
        let record = BoolErrRecord::new(0, 0, 0x0F, 1, 0);
        assert_eq!(record.id(), 0x0205);
        assert_eq!(record.data().len(), 8);
    }

    #[test]
    fn test_row_data_to_cell_records_text() {
        use crate::xls::records::workbook::sst_record::SharedStringTable;

        let row_data = vec![Some(XlsCell::Text("Hello".to_string()))];
        let mut sst = SharedStringTable::new();
        let result = row_data_to_cell_records(0, &row_data, 0x0F, &mut sst);

        assert!(!result.is_empty());
        assert_eq!(&result[0..2], &0x00FDu16.to_le_bytes());
    }

    #[test]
    fn test_row_data_to_cell_records_number() {
        use crate::xls::records::workbook::sst_record::SharedStringTable;

        let row_data = vec![Some(XlsCell::Number(42.0))];
        let mut sst = SharedStringTable::new();
        let result = row_data_to_cell_records(0, &row_data, 0x0F, &mut sst);

        assert!(!result.is_empty());
    }

    #[test]
    fn test_row_data_to_cell_records_boolean() {
        use crate::xls::records::workbook::sst_record::SharedStringTable;

        let row_data = vec![Some(XlsCell::Boolean(true))];
        let mut sst = SharedStringTable::new();
        let result = row_data_to_cell_records(0, &row_data, 0x0F, &mut sst);

        assert!(!result.is_empty());
        assert_eq!(&result[0..2], &0x0205u16.to_le_bytes());
    }

    #[test]
    fn test_row_data_to_cell_records_blank() {
        use crate::xls::records::workbook::sst_record::SharedStringTable;

        let row_data: Vec<Option<XlsCell>> = vec![None, None, None];
        let mut sst = SharedStringTable::new();
        let result = row_data_to_cell_records(0, &row_data, 0x0F, &mut sst);

        assert!(!result.is_empty());
    }
}
