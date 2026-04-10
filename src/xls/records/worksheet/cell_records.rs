use crate::xls::records::workbook::sst_record::SharedStringTable;
use crate::xls::records::{decode_rk_value, BiffRecord, ParsableRecord, ParseState};
use crate::xls::{XlsCell, XlsError};

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
        let count = (self.last_col as i32 - self.first_col as i32 + 1).max(0) as usize;
        let mut buf = Vec::with_capacity(4 + count * 6);
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

/// FormulaRecord 记录（简化版）
///
/// 作用：存储公式单元格及其结果值
///
/// 这是一个简化实现，只读取公式的结果值（8字节浮点数），
/// 不解析公式本身。
#[derive(Debug)]
pub enum FormulaResult {
    Number(f64),
    String(u32), // SST index
}

#[derive(Debug)]
pub struct FormulaRecord {
    row: u16,
    col: u16,
    xf_idx: u16,
    result: FormulaResult,
}

impl FormulaRecord {
    pub fn new(row: u16, col: u16, xf_idx: u16, result: f64) -> Self {
        FormulaRecord {
            row,
            col,
            xf_idx,
            result: FormulaResult::Number(result),
        }
    }
}

impl BiffRecord for FormulaRecord {
    fn id(&self) -> u16 {
        0x0006
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(20);
        buf.extend_from_slice(&self.row.to_le_bytes());
        buf.extend_from_slice(&self.col.to_le_bytes());
        buf.extend_from_slice(&self.xf_idx.to_le_bytes());
        match self.result {
            FormulaResult::Number(val) => {
                buf.extend_from_slice(&val.to_le_bytes());
            }
            FormulaResult::String(sst_idx) => {
                // 字符串结果的格式需要进一步验证
                buf.extend_from_slice(&sst_idx.to_le_bytes());
                buf.extend_from_slice(&[0xFF, 0xFF, 0xFF, 0xFF]); // 标志
            }
        }
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
                } else if count > 1 {
                    // 多个连续的可 RK 编码数字
                    let first_col = col;
                    // 确保 j > 0 避免减法溢出
                    let last_col = if j > 0 { (j - 1) as u16 } else { col };

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

// ============================================================================
// ParsableRecord implementations for reading
// ============================================================================

impl ParsableRecord for NumberRecord {
    const RECORD_ID: u16 = 0x0203;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 14 {
            return Err(XlsError::InvalidFormat(format!(
                "NumberRecord data too short: {} bytes",
                data.len()
            )));
        }

        let row = u16::from_le_bytes([data[0], data[1]]);
        let col = u16::from_le_bytes([data[2], data[3]]);
        let xf_idx = u16::from_le_bytes([data[4], data[5]]);
        let number = f64::from_le_bytes([
            data[6], data[7], data[8], data[9], data[10], data[11], data[12], data[13],
        ]);

        Ok(NumberRecord::new(row, col, xf_idx, number))
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        let sheet = state.current_sheet_mut()?;
        sheet.set_cell(
            self.row as usize,
            self.col as usize,
            XlsCell::Number(self.number),
        );
        Ok(())
    }
}

impl ParsableRecord for RKRecord {
    const RECORD_ID: u16 = 0x027E;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 10 {
            return Err(XlsError::InvalidFormat(format!(
                "RKRecord data too short: {} bytes",
                data.len()
            )));
        }

        let row = u16::from_le_bytes([data[0], data[1]]);
        let col = u16::from_le_bytes([data[2], data[3]]);
        let xf_idx = u16::from_le_bytes([data[4], data[5]]);
        let rk_encoded = i32::from_le_bytes([data[6], data[7], data[8], data[9]]);

        Ok(RKRecord::new(row, col, xf_idx, rk_encoded))
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        let sheet = state.current_sheet_mut()?;
        let value = decode_rk_value(self.rk_encoded);
        sheet.set_cell(self.row as usize, self.col as usize, XlsCell::Number(value));
        Ok(())
    }
}

impl ParsableRecord for BlankRecord {
    const RECORD_ID: u16 = 0x0201;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 6 {
            return Err(XlsError::InvalidFormat(format!(
                "BlankRecord data too short: {} bytes",
                data.len()
            )));
        }

        let row = u16::from_le_bytes([data[0], data[1]]);
        let col = u16::from_le_bytes([data[2], data[3]]);
        let xf_idx = u16::from_le_bytes([data[4], data[5]]);

        Ok(BlankRecord::new(row, col, xf_idx))
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        let sheet = state.current_sheet_mut()?;
        // BlankRecord 对应空值，不存储（保持为 None）
        // 如果需要显式存储空值，可以使用：
        // sheet.set_cell(self.row as usize, self.col as usize, XlsCell::Empty);
        let _ = (sheet, self); // 抑制未使用警告
        Ok(())
    }
}

impl ParsableRecord for LabelSSTRecord {
    const RECORD_ID: u16 = 0x00FD;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 10 {
            return Err(XlsError::InvalidFormat(format!(
                "LabelSSTRecord data too short: {} bytes",
                data.len()
            )));
        }

        let row = u16::from_le_bytes([data[0], data[1]]);
        let col = u16::from_le_bytes([data[2], data[3]]);
        let xf_idx = u16::from_le_bytes([data[4], data[5]]);
        let sst_idx = u32::from_le_bytes([data[6], data[7], data[8], data[9]]);

        Ok(LabelSSTRecord::new(row, col, xf_idx, sst_idx))
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        // 先获取字符串，避免借用冲突
        let text = if self.sst_idx as usize >= state.sst.get_strings().len() {
            // SST 索引越界，使用空字符串并警告
            eprintln!(
                "Warning: SST index {} out of bounds, using empty string",
                self.sst_idx
            );
            String::new()
        } else {
            state.sst.get_strings()[self.sst_idx as usize].clone()
        };

        let sheet = state.current_sheet_mut()?;
        sheet.set_cell(self.row as usize, self.col as usize, XlsCell::Text(text));
        Ok(())
    }
}

impl ParsableRecord for MulRkRecord {
    const RECORD_ID: u16 = 0x00BD;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 6 {
            return Err(XlsError::InvalidFormat(format!(
                "MulRkRecord data too short: {} bytes",
                data.len()
            )));
        }

        let row = u16::from_le_bytes([data[0], data[1]]);
        let first_col = u16::from_le_bytes([data[2], data[3]]);

        // 数据部分：每个 RK 值占 6 字节 (2 字节 xf_idx + 4 字节 rk_encoded)
        // 最后 2 字节是 last_col
        let data_len = data.len();
        if data_len < 6 {
            return Err(XlsError::InvalidFormat(
                "MulRkRecord data too short for last_col".to_string(),
            ));
        }

        let last_col = u16::from_le_bytes([data[data_len - 2], data[data_len - 1]]);

        // 解析中间的 RK 值
        let mut rk_values = Vec::new();
        let mut offset = 4; // 跳过 row 和 first_col

        while offset + 6 <= data_len - 2 {
            let xf_idx = u16::from_le_bytes([data[offset], data[offset + 1]]);
            let rk_encoded = i32::from_le_bytes([
                data[offset + 2],
                data[offset + 3],
                data[offset + 4],
                data[offset + 5],
            ]);
            rk_values.push((xf_idx, rk_encoded));
            offset += 6;
        }

        Ok(MulRkRecord::new(row, first_col, last_col, rk_values))
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        let sheet = state.current_sheet_mut()?;

        // 展开为多个单元格
        for (i, (xf_idx, rk_encoded)) in self.rk_values.iter().enumerate() {
            let col = self.first_col as usize + i;
            let value = decode_rk_value(*rk_encoded);

            // 如果需要，可以存储 xf_idx，但目前只存储值
            let _ = xf_idx; // 抑制未使用警告

            sheet.set_cell(self.row as usize, col, XlsCell::Number(value));
        }

        Ok(())
    }
}

impl ParsableRecord for MulBlankRecord {
    const RECORD_ID: u16 = 0x00BE;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 6 {
            return Err(XlsError::InvalidFormat(format!(
                "MulBlankRecord data too short: {} bytes",
                data.len()
            )));
        }

        let row = u16::from_le_bytes([data[0], data[1]]);
        let first_col = u16::from_le_bytes([data[2], data[3]]);

        // 数据部分：每个空白单元格的 xf_idx 占 2 字节
        // 最后 2 字节是 last_col
        let data_len = data.len();
        let last_col = u16::from_le_bytes([data[data_len - 2], data[data_len - 1]]);

        Ok(MulBlankRecord::new(row, first_col, last_col, 0))
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        let sheet = state.current_sheet_mut()?;

        // Blank 单元格不存储（保持为 None）
        // 或者如果需要显式标记，可以设置为空值
        let _ = sheet; // 抑制未使用警告
        let _ = self;

        Ok(())
    }
}

impl ParsableRecord for BoolErrRecord {
    const RECORD_ID: u16 = 0x0205;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 8 {
            return Err(XlsError::InvalidFormat(format!(
                "BoolErrRecord data too short: {} bytes",
                data.len()
            )));
        }

        let row = u16::from_le_bytes([data[0], data[1]]);
        let col = u16::from_le_bytes([data[2], data[3]]);
        let xf_idx = u16::from_le_bytes([data[4], data[5]]);
        let value = data[6];
        let is_error = data[7];

        Ok(BoolErrRecord::new(row, col, xf_idx, value, is_error))
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        let sheet = state.current_sheet_mut()?;

        // 目前只处理布尔值，错误值当作 None 处理
        if self.is_error == 0 {
            // 布尔值: 0 = FALSE, 1 = TRUE
            let bool_value = self.value != 0;
            sheet.set_cell(
                self.row as usize,
                self.col as usize,
                XlsCell::Boolean(bool_value),
            );
        }
        // 错误值不存储（保持为 None）

        Ok(())
    }
}

impl ParsableRecord for FormulaRecord {
    const RECORD_ID: u16 = 0x0006;

    fn parse(data: &[u8]) -> Result<Self, XlsError> {
        if data.len() < 14 {
            return Err(XlsError::InvalidFormat(format!(
                "FormulaRecord data too short: {} bytes",
                data.len()
            )));
        }

        let row = u16::from_le_bytes([data[0], data[1]]);
        let col = u16::from_le_bytes([data[2], data[3]]);
        let xf_idx = u16::from_le_bytes([data[4], data[5]]);

        let result_bytes = &data[6..14];

        // 检测字符串结果：result[6] == 0xFF && result[7] == 0xFF
        let is_string = result_bytes[6] == 0xFF && result_bytes[7] == 0xFF;

        let result = if is_string {
            // 字符串结果：从前 4 字节读取 SST 索引
            let sst_idx = u32::from_le_bytes([
                result_bytes[0],
                result_bytes[1],
                result_bytes[2],
                result_bytes[3],
            ]);
            FormulaResult::String(sst_idx)
        } else {
            // 数字结果
            let value = f64::from_le_bytes([
                result_bytes[0],
                result_bytes[1],
                result_bytes[2],
                result_bytes[3],
                result_bytes[4],
                result_bytes[5],
                result_bytes[6],
                result_bytes[7],
            ]);
            FormulaResult::Number(value)
        };

        Ok(FormulaRecord {
            row,
            col,
            xf_idx,
            result,
        })
    }

    fn apply(&self, state: &mut ParseState) -> Result<(), XlsError> {
        // 先获取字符串数据（避免借用冲突）
        let text_for_string_result = if let FormulaResult::String(sst_idx) = &self.result {
            if (*sst_idx as usize) < state.sst.get_strings().len() {
                Some(state.sst.get_strings()[*sst_idx as usize].clone())
            } else {
                eprintln!("Warning: Formula SST index {} out of bounds", sst_idx);
                Some(String::new())
            }
        } else {
            None
        };

        let sheet = state.current_sheet_mut()?;

        match &self.result {
            FormulaResult::Number(value) => {
                sheet.set_cell(
                    self.row as usize,
                    self.col as usize,
                    XlsCell::Number(*value),
                );
            }
            FormulaResult::String(_) => {
                if let Some(text) = text_for_string_result {
                    sheet.set_cell(self.row as usize, self.col as usize, XlsCell::Text(text));
                }
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod parsable_tests {
    use super::*;
    use crate::xls::XlsSheet;

    #[test]
    fn test_number_record_parse() {
        // NumberRecord: row=1, col=2, xf_idx=3, number=3.14159
        let data = [
            0x01, 0x00, // row = 1
            0x02, 0x00, // col = 2
            0x03, 0x00, // xf_idx = 3
            0x6E, 0x86, 0x1B, 0x0F, 0xF9, 0x21, 0x09, 0x40, // 3.14159 as f64
        ];

        let record = NumberRecord::parse(&data).unwrap();
        assert_eq!(record.row, 1);
        assert_eq!(record.col, 2);
        assert_eq!(record.xf_idx, 3);
        assert!((record.number - std::f64::consts::PI).abs() < 0.00001);
    }

    #[test]
    fn test_number_record_parse_too_short() {
        let data = [0x01, 0x00, 0x02, 0x00]; // only 4 bytes
        assert!(NumberRecord::parse(&data).is_err());
    }

    #[test]
    fn test_number_record_apply() {
        let record = NumberRecord::new(5, 3, 0, 42.0);
        let mut state = ParseState::new();

        // Create a dummy sheet
        state.current_sheet = Some(XlsSheet::new("Test".to_string()));

        record.apply(&mut state).unwrap();

        // Verify the cell was set
        let sheet = state.current_sheet.unwrap();
        match &sheet.rows[5][3] {
            Some(XlsCell::Number(n)) => assert_eq!(*n, 42.0),
            _ => panic!("Expected Number cell"),
        }
    }

    #[test]
    fn test_rk_record_parse() {
        // RKRecord: row=0, col=1, xf_idx=15, rk_encoded=170 (42 << 2 | 0x02)
        let data = [
            0x00, 0x00, // row = 0
            0x01, 0x00, // col = 1
            0x0F, 0x00, // xf_idx = 15
            0xAA, 0x00, 0x00, 0x00, // rk_encoded = 170
        ];

        let record = RKRecord::parse(&data).unwrap();
        assert_eq!(record.row, 0);
        assert_eq!(record.col, 1);
        assert_eq!(record.xf_idx, 15);
        assert_eq!(record.rk_encoded, 170);
    }

    #[test]
    fn test_rk_record_apply() {
        // RK encoded integer 42: (42 << 2) | 0x02 = 170
        let record = RKRecord::new(2, 4, 0, 170);
        let mut state = ParseState::new();

        state.current_sheet = Some(XlsSheet::new("Test".to_string()));

        record.apply(&mut state).unwrap();

        let sheet = state.current_sheet.unwrap();
        match &sheet.rows[2][4] {
            Some(XlsCell::Number(n)) => assert_eq!(*n, 42.0),
            _ => panic!("Expected Number cell with value 42.0"),
        }
    }

    #[test]
    fn test_blank_record_parse() {
        let data = [
            0x0A, 0x00, // row = 10
            0x05, 0x00, // col = 5
            0x0F, 0x00, // xf_idx = 15
        ];

        let record = BlankRecord::parse(&data).unwrap();
        assert_eq!(record.row, 10);
        assert_eq!(record.col, 5);
        assert_eq!(record.xf_idx, 15);
    }

    #[test]
    fn test_blank_record_apply() {
        let record = BlankRecord::new(7, 8, 0);
        let mut state = ParseState::new();

        // Pre-populate with some data that should be cleared
        let mut sheet = XlsSheet::new("Test".to_string());
        sheet.set_cell(7, 8, XlsCell::Number(999.0));
        state.current_sheet = Some(sheet);

        // BlankRecord should not modify the cell (leave as is or empty)
        record.apply(&mut state).unwrap();

        // Verify cell is still there (BlankRecord doesn't clear existing data)
        let sheet = state.current_sheet.unwrap();
        assert!(sheet.rows[7][8].is_some());
    }

    #[test]
    fn test_label_sst_record_parse() {
        let data = [
            0x03, 0x00, // row = 3
            0x02, 0x00, // col = 2
            0x0F, 0x00, // xf_idx = 15
            0x00, 0x00, 0x00, 0x00, // sst_idx = 0
        ];

        let record = LabelSSTRecord::parse(&data).unwrap();
        assert_eq!(record.row, 3);
        assert_eq!(record.col, 2);
        assert_eq!(record.xf_idx, 15);
        assert_eq!(record.sst_idx, 0);
    }

    #[test]
    fn test_label_sst_record_apply() {
        let record = LabelSSTRecord::new(1, 2, 0, 0);
        let mut state = ParseState::new();

        // Add a string to SST
        state.sst.add("Hello World".to_string());

        state.current_sheet = Some(XlsSheet::new("Test".to_string()));

        record.apply(&mut state).unwrap();

        let sheet = state.current_sheet.unwrap();
        match &sheet.rows[1][2] {
            Some(XlsCell::Text(t)) => assert_eq!(t, "Hello World"),
            _ => panic!("Expected Text cell"),
        }
    }

    #[test]
    fn test_label_sst_record_apply_out_of_bounds() {
        let record = LabelSSTRecord::new(0, 0, 0, 999); // Invalid SST index
        let mut state = ParseState::new();

        state.current_sheet = Some(XlsSheet::new("Test".to_string()));

        // Should not panic, uses empty string
        record.apply(&mut state).unwrap();

        let sheet = state.current_sheet.unwrap();
        match &sheet.rows[0][0] {
            Some(XlsCell::Text(t)) => assert!(t.is_empty()),
            _ => panic!("Expected empty Text cell"),
        }
    }
}
