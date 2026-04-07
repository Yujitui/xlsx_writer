use byteorder::{LittleEndian, ReadBytesExt};
use cfb::CompoundFile;
use encoding_rs::UTF_16LE;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek, SeekFrom};
use crate::xls_reader::{XlsError, XlsRecordReader};

// 工作簿/工作表结构相关
const RT_BOF        : u16 = 0x0809; // 文件开始标记 (Beginning of File)
const RT_BOUNDSHEET : u16 = 0x0085; // 工作表信息 (Sheet Information)
const RT_SST        : u16 = 0x00FC; // 共享字符串表 (Shared String Table)
const RT_DIMENSIONS : u16 = 0x0200; // 工作表维度信息 (Sheet Dimensions)
const RT_ROW        : u16 = 0x0208; // 行信息 (Row Information)
const RT_EOF        : u16 = 0x000A; // 文件结束标记 (End of File)

// 单元格数据类型
const RT_NUMBER     : u16 = 0x0203; // 数值单元格 (Floating Point Number)
const RT_RK         : u16 = 0x027E; // 编码后的整数/浮点数 (Encoded Integer/Float)
const RT_RK_OLD     : u16 = 0x007E; // 旧版 RK 记录 (Older version of RK)
const RT_MULRK      : u16 = 0x00BD; // 连续多个 RK 数值 (Multiple RK Records)
const RT_LABEL      : u16 = 0x0204; // 内联字符串标签 (In-place String Label, 已弃用)
const RT_LABELSST   : u16 = 0x00FD; // 来自 SST 的字符串标签 (String Label from SST)
const RT_FORMULA    : u16 = 0x0006; // 公式单元格

/// 代表 Excel 工作表中的一个单元格。
///
/// 此枚举封装了单元格可以包含的不同数据类型。
/// 每个变体对应于 Excel 单元格中存在的特定数据类型。
///
/// # 变体 (Variants)
///
/// * `Number(f64)` - 一个数值。Excel 内部将数字存储为 64 位浮点数。
///                   这包括整数、小数以及日期（以序列号形式存储）。
/// * `Text(String)` - 一串字符。这可以是纯文本，也可以是公式计算结果为文本的情况。
/// * `Boolean(bool)` - 一个逻辑值，`true` 或 `false`。
#[derive(Debug, Clone, PartialEq)]
pub enum XlsCell {
    /// 表示存储为 64 位浮点数的数值。
    /// 这是 Excel 中数字的标准内部表示形式。
    Number(f64),

    /// 表示存储为 Rust `String` 类型的文本数据。
    /// 这包括标签、名称以及任何公式得出的结果文本。
    Text(String),

    /// 表示一个布尔逻辑值（`true` 或 `false`）。
    Boolean(bool),
}

/// 代表一个 Excel 工作表。
///
/// 此结构体封装了单个工作表的所有信息，包括其名称和包含的单元格数据。
/// 数据以二维向量的形式存储，模拟了工作表的行列结构。
/// `rows` 向量的每个元素代表一行，行内的 `Vec<Option<XlsCell>>` 代表该行的各个单元格。
/// 使用 `Option` 是因为并非所有行列交叉点都一定有数据，`None` 表示该单元格为空。
#[derive(Debug)]
pub struct XlsSheet {
    /// 工作表的名称，例如 "Sheet1" 或用户自定义的名称。
    pub sheet_name: String,
    /// 存储工作表所有数据的二维向量。
    /// 外层向量的索引对应行号（从 0 开始）。
    /// 内层向量的索引对应列号（从 0 开始）。
    /// `Option<XlsCell>` 用于表示单元格是否有值：
    /// - `Some(cell)` 表示该单元格包含数据。
    /// - `None` 表示该单元格是空的。
    pub rows: Vec<Vec<Option<XlsCell>>>,
}

impl XlsSheet {
    /// 获取工作表中包含数据的实际范围。
    ///
    /// 此函数分析 `rows` 向量，找出包含至少一个非空单元格的最后行和最后列，
    /// 从而确定有效数据区域的边界。
    ///
    /// # 返回值
    ///
    /// * `Some((max_row, max_col))` - 如果工作表包含数据，则返回一个元组，
    ///   其中 `max_row` 是最后一行的索引（从 0 开始），
    ///   `max_col` 是最后一列的索引（从 0 开始）。
    /// * `None` - 如果工作表没有任何数据（即 `rows` 为空或所有单元格都为空）。
    pub fn data_range(&self) -> Option<(usize, usize)> {
        if self.rows.is_empty() {
            return None;
        }

        let max_row = self.rows.len() - 1;
        let max_col = self.rows.iter()
            .map(|row| row.len())
            .max()
            .unwrap_or(0);

        if max_col == 0 {
            None
        } else {
            Some((max_row, max_col - 1)) // 0-indexed
        }
    }

    /// 为遍历工作表中的所有非空单元格提供一个迭代器。
    ///
    /// 此函数返回一个迭代器，该迭代器会按行优先的顺序（从上到下，从左到右）
    /// 产出每一个包含数据的单元格及其对应的行列坐标。
    ///
    /// # 返回值
    ///
    /// 返回一个迭代器，其项目（Item）类型为 `(usize, usize, &XlsCell)`：
    /// - 第一个 `usize` 是行索引（从 0 开始）。
    /// - 第二个 `usize` 是列索引（从 0 开始）。
    /// - `&XlsCell` 是指向单元格数据的引用。
    pub fn cell_iterator(&self) -> impl Iterator<Item = (usize, usize, &XlsCell)> {
        self.rows.iter().enumerate().flat_map(|(row_idx, row)| {
            row.iter().enumerate().filter_map(move |(col_idx, cell)| {
                cell.as_ref().map(|c| (row_idx, col_idx, c))
            })
        })
    }

    /// 安全地在指定行列位置设置单元格值
    ///
    /// # 参数
    /// * `row` - 行索引（从 0 开始）
    /// * `col` - 列索引（从 0 开始）
    /// * `cell` - 要设置的单元格内容
    fn set_cell(&mut self, row: usize, col: usize, cell: XlsCell) {
        // 自动扩展行数（如果需要）
        if self.rows.len() <= row {
            self.rows.resize_with(row + 1, || vec![]);
        }

        // 自动扩展列数（如果需要）
        if self.rows[row].len() <= col {
            self.rows[row].resize_with(col + 1, || None);
        }

        // 设置单元格值
        self.rows[row][col] = Some(cell);
    }
}

/// Excel (.xls) 文件读取器。
///
/// 此结构体是解析和访问 Excel 97-2003 工作簿（.xls 格式）的核心。
/// 它负责打开文件，解析其内部的二进制结构（如 CFB/OLE2），
/// 并将数据提取为易于操作的 `XlsSheet` 对象集合。
///
/// 一个 `XlsReader` 实例通常代表整个工作簿，并包含其中所有的 `XlsSheet`。
#[derive(Debug)]
pub struct XlsReader {
    /// 存储工作簿中所有工作表的向量。
    ///
    /// 每个工作表都是一个 `XlsSheet` 结构体实例，包含了该工作表的名称和数据。
    /// 向量的顺序通常反映了工作表在原始 Excel 文件中的顺序。
    /// 可以通过索引（0-based）、名称等方式访问这些工作表。
    pub sheets: Vec<XlsSheet>,
}

impl XlsReader {
    /// 创建一个新的 `XlsReader` 实例。
    ///
    /// 这是一个构造函数，用于初始化一个 `XlsReader` 对象。
    ///
    /// # 返回值
    ///
    /// 返回一个空的 `XlsReader` 新实例。
    pub fn new() -> Self {
        XlsReader { sheets: Vec::new() }
    }

    /// 获取工作簿中工作表的数量。
    ///
    /// # 返回值
    ///
    /// 返回 `sheets` 向量的长度，即工作簿中包含的工作表总数。
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// 根据索引获取工作表的不可变引用。
    ///
    /// 索引是从 0 开始的。例如，索引 0 对应第一个工作表。
    ///
    /// # 参数
    ///
    /// * `index`: 工作表的从零开始的索引。
    ///
    /// # 返回值
    ///
    /// * `Some(&XlsSheet)`: 如果索引有效，则返回指向该工作表的引用。
    /// * `None`: 如果索引超出了工作表列表的范围。
    pub fn sheet_at(&self, index: usize) -> Option<&XlsSheet> {
        self.sheets.get(index)
    }

    /// 根据名称获取工作表的不可变引用。
    ///
    /// 此方法会遍历所有工作表，查找名称完全匹配的工作表。
    ///
    /// # 参数
    ///
    /// * `name`: 要查找的工作表名称。
    ///
    /// # 返回值
    ///
    /// * `Some(&XlsSheet)`: 如果找到名称匹配的工作表，则返回指向该工作表的引用。
    /// * `None`: 如果没有找到具有给定名称的工作表。
    pub fn sheet_by_name(&self, name: &str) -> Option<&XlsSheet> {
        self.sheets.iter().find(|sheet| sheet.sheet_name == name)
    }

    /// 获取工作簿中所有工作表的名称。
    ///
    /// # 返回值
    ///
    /// 返回一个包含所有工作表名称的 `Vec<&String>`。
    /// 向量中名称的顺序与它们在 `sheets` 向量中的顺序一致。
    pub fn sheet_names(&self) -> Vec<&String> {
        self.sheets.iter().map(|sheet| &sheet.sheet_name).collect()
    }

    /// 获取工作簿中的第一个工作表的不可变引用。
    ///
    /// 这通常是 Excel 默认激活或显示的工作表。
    ///
    /// # 返回值
    ///
    /// * `Some(&XlsSheet)`: 如果工作簿中至少有一个工作表，则返回第一个工作表的引用。
    /// * `None`: 如果工作簿中没有任何工作表（即 `sheets` 为空）。
    pub fn first_sheet(&self) -> Option<&XlsSheet> {
        self.sheets.first()
    }

    /// 检查工作簿中是否存在具有指定名称的工作表。
    ///
    /// # 参数
    ///
    /// * `name`: 要检查的工作表名称。
    ///
    /// # 返回值
    ///
    /// * `true`: 如果存在具有该名称的工作表。
    /// * `false`: 如果不存在具有该名称的工作表。
    pub fn contains_sheet(&self, name: &str) -> bool {
        self.sheets.iter().any(|sheet| sheet.sheet_name == name)
    }
}

impl XlsReader {

    /// 从指定路径读取并解析一个 Excel (.xls) 文件。
    ///
    /// 此函数是 `XlsReader` 的核心公共接口，负责执行整个文件加载和解析过程。
    /// 它会打开文件，处理其复合文件结构（CFB/OLE2），定位到包含工作表数据的 "Workbook" 流，
    /// 然后调用内部解析函数来提取数据，并将结果存储在 `self.sheets` 中。
    ///
    /// # 参数说明
    ///
    /// * `mut self`: 该方法会消耗（consume）当前的 `XlsReader` 实例。
    ///               这意味着调用此方法后，原来的实例将不再可用。
    ///               这是一种常见的模式，特别是在构建或修改对象时。
    /// * `path`: 一个字符串切片（`&str`），指向要读取的 .xls 文件的磁盘路径。
    ///
    /// # 返回值说明
    ///
    /// * 成功时 (`Ok(Self)`): 返回一个 `XlsReader` 实例，其 `sheets` 字段已被填充
    ///                        为从文件中解析出的所有工作表。
    /// * 失败时 (`Err(XlsError)`): 返回一个 `XlsError` 枚举实例，描述了在读取或解析过程中
    ///                             发生的具体错误类型，例如文件IO错误、格式无效或意外的文件结尾等。
    pub fn read_xls(mut self, path: &str) -> Result<Self, XlsError> {
        // 打开文件
        let file = File::open(path).map_err(XlsError::IoError)?;
        let mut buf_reader = BufReader::new(file);
        let mut buffer = Vec::new();
        buf_reader.read_to_end(&mut buffer).map_err(XlsError::IoError)?;

        // 将文件内容作为 Cursor 处理，便于随机访问
        let mut cursor = Cursor::new(buffer);

        // 打开 CFB 结构（注意这里可能出错）
        let mut cfb = CompoundFile::open(&mut cursor).map_err(XlsError::IoError)?;

        // 解析 Workbook 流
        let workbook_stream = cfb.open_stream("Workbook")
            .map_err(|_| XlsError::InvalidFormat("Workbook stream not found".into()))?;

        // 将流中的所有字节收集到一个 Vec<u8> 中，以便后续解析。
        let bytes: Vec<u8> = workbook_stream.bytes().collect::<Result<Vec<_>, _>>()
            .map_err(XlsError::IoError)?;

        // 创建一个新的 Cursor 来包装这些字节，供 `parse_workbook` 方法读取。
        let mut reader = Cursor::new(bytes);

        match Self::parse_workbook(&mut reader) {
            Ok(sheets) => {
                // 解析成功，将解析出的所有工作表添加到当前实例中。
                for sheet in sheets {
                    if !self.contains_sheet(sheet.sheet_name.as_str()) {
                        // 如果名称不冲突，则将工作表添加到 `self.sheets` 向量的末尾。
                        self.sheets.push(sheet);
                    } else {
                        eprintln!("sheet '{}' 与现有表格重复", sheet.sheet_name);
                        continue;
                    }
                }
            },
            Err(e) => return Err(e),
        }

        Ok(self)
    }

    fn parse_bof<R: Read + Seek>(
        reader: &mut R, current_sheet: &mut Option<XlsSheet>,
        sheets: &Vec<XlsSheet>, sheet_names: &Vec<String>,
    ) -> Result<(), XlsError> {
        let _version = reader.read_u16::<LittleEndian>()?;
        let substream_type = reader.read_u16::<LittleEndian>()?;

        match substream_type {
            0x0005 => println!("正在解析：全局工作簿信息 (Workbook Globals)"),
            0x0010 => {
                let name = sheet_names.get(sheets.len())
                    .cloned()
                    .unwrap_or_else(|| format!("Sheet{}", sheets.len() + 1));
                *current_sheet = Some(XlsSheet { sheet_name: name, rows: Vec::new() });
            },
            _ => {},
        }
        Ok(())
    }

    fn parse_bound_sheet<R: Read + Seek>(
        reader: &mut R, sheet_names: &mut Vec<String>,
    ) -> Result<(), XlsError> {
        let _pos = reader.read_u32::<LittleEndian>()?;
        let _visibility = reader.read_u8()?;
        let _type = reader.read_u8()?;

        let char_count = reader.read_u8()? as usize; // 字符数
        let grbit = reader.read_u8()?;               // 编码标志位

        let is_utf16 = (grbit & 0x01) != 0;

        let sheet_name = if is_utf16 {
            // UTF-16LE 模式：每个字符 2 字节
            let mut buf = vec![0u8; char_count * 2];
            reader.read_exact(&mut buf)?;
            let (cow, _, _) = UTF_16LE.decode(&buf);
            cow.into_owned()
        } else {
            // Latin-1 模式：每个字符 1 字节
            let mut buf = vec![0u8; char_count];
            reader.read_exact(&mut buf)?;
            // 即使是 Latin-1，也建议通过 lossy 转换确保安全
            String::from_utf8_lossy(&buf).to_string()
        };

        sheet_names.push(sheet_name);
        Ok(())
    }

    fn parse_sst<R: Read + Seek>(
        reader: &mut R, length: &usize,
        sst: &mut Vec<String>
    ) -> Result<(), XlsError> {
        let mut rec_reader = XlsRecordReader::new(reader, *length);
        let _total_strings = rec_reader.inner.read_u32::<LittleEndian>()?; // 全局总数
        let unique_count = rec_reader.inner.read_u32::<LittleEndian>()?;   // 唯一数
        rec_reader.current_record_remaining -= 8;

        for _ in 0..unique_count {
            let char_count = rec_reader.inner.read_u16::<LittleEndian>()? as usize;
            rec_reader.current_record_remaining -= 2;

            let flag = rec_reader.read_u8()?;
            let is_utf16 = (flag & 0x01) != 0;
            let has_rich = (flag & 0x08) != 0;
            let has_ext = (flag & 0x04) != 0;

            let mut rich_runs = 0;
            if has_rich {
                rich_runs = rec_reader.inner.read_u16::<LittleEndian>()?;
                rec_reader.current_record_remaining -= 2;
            }
            let mut ext_size = 0;
            if has_ext {
                ext_size = rec_reader.inner.read_u32::<LittleEndian>()?;
                rec_reader.current_record_remaining -= 4;
            }

            // 读取字符串主体（这里是最坑的地方，Continue 可能在这里切断）
            let mut string_data = Vec::new();
            let mut remaining_chars = char_count;
            let mut current_is_utf16 = is_utf16;

            while remaining_chars > 0 {
                // 如果遇到 Continue 记录，Continue 开头会有一个新的 flag 字节覆盖编码格式！
                if rec_reader.current_record_remaining == 0 {
                    if !rec_reader.ensure_data()? { break; }
                    let new_flag = rec_reader.read_u8()?;
                    current_is_utf16 = (new_flag & 0x01) != 0;
                }

                if current_is_utf16 {
                    let mut buf = vec![0u8; 2];
                    rec_reader.read_exact(&mut buf)?;
                    string_data.extend_from_slice(&buf);
                } else {
                    let mut buf = vec![0u8; 1];
                    rec_reader.read_exact(&mut buf)?;
                    // 如果后续需要转 UTF16 处理，这里可以补 0 字节
                    string_data.push(buf[0]);
                    string_data.push(0x00);
                }
                remaining_chars -= 1;
            }

            // 最后解码 string_data 为 UTF8
            let (res, _, _) = UTF_16LE.decode(&string_data);
            sst.push(res.into_owned());

            // 跳过 Rich Text 和 Extension 数据
            for _ in 0..(rich_runs * 4) { rec_reader.read_u8()?; }
            for _ in 0..ext_size { rec_reader.read_u8()?; }
        }
        Ok(())
    }

    fn parse_dimensions<R: Read + Seek>(
        reader: &mut R, current_sheet: &mut Option<XlsSheet>,
    ) -> Result<(), XlsError> {
        // 初始化当前工作表的行结构
        let _first_row = reader.read_u32::<LittleEndian>()?;
        let last_row = reader.read_u32::<LittleEndian>()?;
        let _first_col = reader.read_u16::<LittleEndian>()?;
        let last_col = reader.read_u16::<LittleEndian>()?;

        if let Some(ref mut sheet) = current_sheet {
            // 创建新的工作表并预分配空间
            for _ in 0..=last_row {
                sheet.rows.push(vec![None; (last_col + 1) as usize]);
            }
        }

        Ok(())
    }

    fn parse_row<R: Read + Seek>(
        reader: &mut R, current_sheet: &mut Option<XlsSheet>,
    ) -> Result<(), XlsError> {
        let row_index = reader.read_u32::<LittleEndian>()? as usize;
        let _first_col = reader.read_u16::<LittleEndian>()? as usize;
        let last_col = reader.read_u16::<LittleEndian>()? as usize;

        // 确保行存在并初始化列
        if let Some(ref mut sheet) = current_sheet {
            if sheet.rows.len() <= row_index {
                sheet.rows.resize_with(row_index + 1, || vec![]);
            }
            sheet.rows[row_index].resize_with(last_col + 1, || None);
        }
        Ok(())
    }

    fn parse_number<R: Read + Seek>(
        reader: &mut R, current_sheet: &mut Option<XlsSheet>,
    ) -> Result<(), XlsError> {
        let row = reader.read_u16::<LittleEndian>()? as usize;
        let col = reader.read_u16::<LittleEndian>()? as usize;
        let _xf_index = reader.read_u16::<LittleEndian>()?;
        let num = reader.read_f64::<LittleEndian>()?;

        if let Some(ref mut sheet) = current_sheet {
            sheet.set_cell(row, col, XlsCell::Number(num));
        }
        Ok(())
    }

    fn parse_rk<R: Read + Seek>(
        reader: &mut R, current_sheet: &mut Option<XlsSheet>,
    ) -> Result<(), XlsError> {
        let row = reader.read_u16::<LittleEndian>()? as usize;
        let col = reader.read_u16::<LittleEndian>()? as usize;
        let _xf_index = reader.read_u16::<LittleEndian>()?;
        let rk_raw = reader.read_u32::<LittleEndian>()?; // 读取 4 字节原始 RK 值

        // --- RK 解码算法 ---
        let is_multiplied_by_100 = (rk_raw & 0x01) != 0;
        let is_integer = (rk_raw & 0x02) != 0;

        let mut val: f64;

        if is_integer {
            // 情况 A: 这是一个被左移 2 位的带符号整数
            // 使用 i32 保持符号位，右移 2 位还原
            let int_val = (rk_raw as i32) >> 2;
            val = int_val as f64;
        } else {
            // 情况 B: 这是一个 IEEE 754 浮点数的高 30 位
            // 低 32 位被抹零了（因为它是压缩存储）
            let f_bits = (rk_raw as u64 & 0xFFFF_FFFC_u64) << 32;
            val = f64::from_bits(f_bits);
        }

        // 如果标志位说要除以 100（例如存储的是 12345 代表 123.45）
        if is_multiplied_by_100 {
            val /= 100.0;
        }

        // 将结果存入你的 Sheet
        if let Some(ref mut sheet) = current_sheet {
            sheet.set_cell(row, col, XlsCell::Number(val));
        }
        Ok(())
    }

    fn parse_mulrk<R: Read + Seek>(
        reader: &mut R, current_sheet: &mut Option<XlsSheet>,
        length: &usize,
    ) -> Result<(), XlsError> {
        let row = reader.read_u16::<LittleEndian>()? as usize;
        let first_col = reader.read_u16::<LittleEndian>()? as usize;

        // 计算中间 RK 列表的字节数
        // 总长度 - 2(row) - 2(first_col) - 2(last_col)
        let rk_list_len = (length - 6) / 6;

        for i in 0..rk_list_len {
            let _xf_index = reader.read_u16::<LittleEndian>()?;
            let rk_raw = reader.read_u32::<LittleEndian>()?;

            // 复用之前的 RK 解码逻辑
            let is_div_100 = (rk_raw & 0x01) != 0;
            let is_int = (rk_raw & 0x02) != 0;

            let mut val: f64;
            if is_int {
                val = ((rk_raw as i32) >> 2) as f64;
            } else {
                let f_bits = (rk_raw as u64 & 0xFFFF_FFFC_u64) << 32;
                val = f64::from_bits(f_bits);
            }

            if is_div_100 {
                val /= 100.0;
            }

            // 计算当前单元格的列索引
            let col = first_col + i;

            if let Some(ref mut sheet) = current_sheet {
                sheet.set_cell(row, col, XlsCell::Number(val));
            }
        }

        // 最后读取结束列号（校验用，通常直接读掉即可）
        let _last_col = reader.read_u16::<LittleEndian>()?;
        Ok(())
    }

    fn parse_label<R: Read + Seek>(
        reader: &mut R, current_sheet: &mut Option<XlsSheet>,
    ) -> Result<(), XlsError> {
        let row = reader.read_u16::<LittleEndian>()? as usize;
        let col = reader.read_u16::<LittleEndian>()? as usize;
        let _xf_index = reader.read_u16::<LittleEndian>()?;
        let str_len = reader.read_u16::<LittleEndian>()? as usize; // 字符串长度
        let mut buf = vec![0u8; str_len];
        reader.read_exact(&mut buf)?; // 读取字符串字节
        let text = String::from_utf8_lossy(&buf).to_string(); // 转换为字符串

        if let Some(ref mut sheet) = current_sheet {
            sheet.set_cell(row, col, XlsCell::Text(text));
        }
        Ok(())
    }

    fn parse_label_sst<R: Read + Seek>(
        reader: &mut R, current_sheet: &mut Option<XlsSheet>,
        sst: &Vec<String>,
    ) -> Result<(), XlsError> {
        let row = reader.read_u16::<LittleEndian>()? as usize;
        let col = reader.read_u16::<LittleEndian>()? as usize;
        let _xf_index = reader.read_u16::<LittleEndian>()?;
        let sst_index = reader.read_u32::<LittleEndian>()? as usize; // 注意这里是 u32

        if !sst.is_empty() && sst_index < sst.len() {
            let text = sst[sst_index].clone();
            if let Some(ref mut sheet) = current_sheet {
                sheet.set_cell(row, col, XlsCell::Text(text));
            }
        }
        Ok(())
    }

    fn parse_formula<R: Read + Seek>(
        reader: &mut R, current_sheet: &mut Option<XlsSheet>,
        length: &usize,
    ) -> Result<(), XlsError> {
        let row = reader.read_u16::<LittleEndian>()? as usize;
        let col = reader.read_u16::<LittleEndian>()? as usize;
        let _xf_index = reader.read_u16::<LittleEndian>()?;

        // 读取 IEEE 浮点数形式的结果值（固定8字节）
        let result_val = reader.read_f64::<LittleEndian>()?;

        // 后续还有其他字段（如公式长度等），但我们只关心结果值即可
        // 可跳过剩余部分：length - 14 字节（前面共读取了 2+2+2+8=14 字节）
        reader.seek(SeekFrom::Current((length - 14) as i64))?;

        if let Some(ref mut sheet) = current_sheet {
            if sheet.rows.len() > row && sheet.rows[row].len() > col {
                sheet.rows[row][col] = Some(XlsCell::Number(result_val));
            }
        }
        Ok(())
    }

    fn parse_eof(
        current_sheet: &mut Option<XlsSheet>, sheets: &mut Vec<XlsSheet>,
    ) -> Result<(), XlsError> {
        if let Some(sheet) = current_sheet.take() {
            sheets.push(sheet); // 将完成的工作表加入列表
        }
        // 如果所有预期的工作表都已处理完毕
        // if sheets.len() >= sheet_names.len() {
        //     break;
        // }
        Ok(())
    }

    fn parse_workbook<R: Read + Seek>(
        reader: &mut R,
    ) -> Result<Vec<XlsSheet>, XlsError> {
        let mut sheets = Vec::new();                    // 存储所有工作表
        let mut current_sheet: Option<XlsSheet> = None; // 当前正在解析的工作表
        let mut sheet_names: Vec<String> = Vec::new();
        let mut sst: Vec<String> = Vec::new();

        loop {
            // 读取记录类型和长度（每个记录头占 4 字节）
            let record_type = match reader.read_u16::<LittleEndian>() {
                Ok(v) => v,
                Err(_) => {
                    // 文件读取完毕，保存最后一个工作表（如果存在）
                    if let Some(sheet) = current_sheet.take() {
                        sheets.push(sheet);
                    }
                    break;
                }
            };

            let length = match reader.read_u16::<LittleEndian>() {
                Ok(v) => v as usize,
                Err(_) => {
                    if let Some(sheet) = current_sheet.take() {
                        sheets.push(sheet);
                    }
                    break;
                }
            };

            // 记录当前 Record 数据的起始位置，用于最后校验或跳过
            let start_pos = reader.stream_position()?;

            match record_type {
                // BOF (Beginning of File) - 表示新部分的开始
                RT_BOF  => Self::parse_bof(reader, &mut current_sheet, &mut sheets, &sheet_names)?,
                // BOUNDSHEET - 定义工作表的位置和名称
                RT_BOUNDSHEET => Self::parse_bound_sheet(reader, &mut sheet_names)?,
                // SST - 定义共享字符串
                RT_SST => Self::parse_sst(reader, &length, &mut sst)?,
                // DIMENSIONS - 指定工作表使用的最大行列范围
                RT_DIMENSIONS => Self::parse_dimensions(reader, &mut current_sheet)?,
                // ROW - 描述单行的属性（如列跨度）
                RT_ROW => Self::parse_row(reader, &mut current_sheet)?,
                // NUMBER - 真正的 8 字节浮点数记录
                RT_NUMBER => Self::parse_number(reader, &mut current_sheet)?,
                // RK (0x027E / 0x007E) - XLS 最常用的数值格式（整型或压缩浮点）
                RT_RK | RT_RK_OLD => Self::parse_rk(reader, &mut current_sheet)?,
                // MULRK (0x00BD) - 一行中连续的多个数字单元格
                RT_MULRK => Self::parse_mulrk(reader, &mut current_sheet, &length)?,
                // LABEL - 内联字符串（已弃用但仍被支持）
                RT_LABEL => Self::parse_label(reader, &mut current_sheet)?,
                // SST - 共享字符串
                RT_LABELSST => Self::parse_label_sst(reader, &mut current_sheet, &sst)?,
                // FORMULA - 公式
                RT_FORMULA => Self::parse_formula(reader, &mut current_sheet, &length)?,
                // EOF (End of File) - 表示当前部分结束
                RT_EOF => Self::parse_eof(&mut current_sheet, &mut sheets)?,
                // 忽略未知或未处理的记录
                _ => { reader.seek(SeekFrom::Current(length as i64))?; } // 跳过该记录内容
            }

            let end_pos = start_pos + length as u64;
            reader.seek(SeekFrom::Start(end_pos))?;
        }

        Ok(sheets) // 返回解析结果
    }
}


#[test]
fn read_test() {
    use super::XlsReader;

    let xls = XlsReader::new()
        .read_xls("/Users/yujitui/Downloads/工作簿1.xls").unwrap();

    for row in xls.first_sheet().unwrap().rows.clone() { // 如果用了 BTreeMap
        let row_str: Vec<&Option<XlsCell>> = row.iter()
            .filter(|x| x.is_some())
            .collect::<Vec<_>>();
        println!("{:?}", row_str);
    }
}