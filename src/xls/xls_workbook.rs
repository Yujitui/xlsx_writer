use crate::xls::{XlsSheet, XlsError, XlsCell, XlsRecordReader, RecordType};
use std::fs::File;
use std::io::{Write, Seek, SeekFrom, BufReader, Cursor, Read};
use cfb::CompoundFile;
use byteorder::{LittleEndian, ReadBytesExt, WriteBytesExt};
use encoding_rs::UTF_16LE;

/// 代表一个 Excel 97-2003 工作簿 (.xls 文件)。
/// 该结构体同时支持读取和写入操作。
pub struct XlsWorkbook {
    /// 存储所有工作表的集合。
    /// 使用 Vec 保持工作表的顺序。
    sheets: Vec<XlsSheet>,

    // （可选）存储文件的其他元数据，如作者、创建时间等。
    // 这里可以根据需要扩展。
}

impl XlsWorkbook {
    /// 创建一个新的 `XlsReader` 实例。
    ///
    /// 这是一个构造函数，用于初始化一个 `XlsReader` 对象。
    ///
    /// # 返回值
    ///
    /// 返回一个空的 `XlsReader` 新实例。
    pub fn new() -> Self {
        XlsWorkbook { sheets: Vec::new() }
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

impl XlsWorkbook {

    /// 将当前解析出的工作簿数据写入到一个新的 .xls 文件中
    pub fn write_xls(&self, output_path: &str) -> Result<(), XlsError> {
        // 创建临时缓冲区用于构建 workbook 流
        let mut workbook_stream = Vec::new();

        // 写入 BOF（Workbook Globals）
        self.write_bof(&mut workbook_stream, true)?;
        // self.write_window1(&mut workbook_stream)?;
        // self.write_default_font(&mut workbook_stream)?;
        // for _ in 0..3 {
        //     self.write_default_xf(&mut workbook_stream)?;
        // }

        // 记录 BoundSheet 起始位置
        let boundsheet_start_pos = workbook_stream.len();

        // 先计算所有 BoundSheet 的实际大小
        let mut boundsheet_total_size = 0;
        for sheet in &self.sheets {
            let name_len = sheet.sheet_name.as_bytes().len();
            let bound_sheet_size = 8 + name_len; // 8字节头部 + 1字节长度 + 名称长度
            boundsheet_total_size += 4 + bound_sheet_size; // Record header (4 bytes) + data
        }

        // 预分配准确的 BoundSheet 空间
        workbook_stream.resize(workbook_stream.len() + boundsheet_total_size, 0);

        // 构建并写入 SST（共享字符串表）
        let sst_strings: Vec<String> = self.collect_all_strings();
        self.write_sst(&mut workbook_stream, &sst_strings)?;

        // Step 2: 构建各个 sheet 的数据流及其偏移映射
        let mut sheet_streams = Vec::new();
        for sheet in &self.sheets {
            let mut sheet_data = Vec::new();
            self.write_sheet_data(&mut sheet_data, sheet, &sst_strings)?;
            sheet_streams.push(sheet_data);
        }

        // Step 3: 写入每个 sheet 数据流，并记录偏移
        let mut sheet_offsets = Vec::new();
        for sheet_data in sheet_streams {
            sheet_offsets.push(workbook_stream.len() as u32); // 当前流的位置即为此 sheet 的偏移
            workbook_stream.extend_from_slice(&sheet_data);
        }

        // 回头更新 BoundSheet 记录
        let mut boundsheet_updated = Vec::new();
        for (i, sheet) in self.sheets.iter().enumerate() {
            let offset = sheet_offsets[i];
            self.write_boundsheet_with_offset(&mut boundsheet_updated, &sheet.sheet_name, offset)?;
        }

        // 精确替换占位符区域
        let boundsheet_actual_end = boundsheet_start_pos + boundsheet_updated.len();
        if boundsheet_actual_end <= workbook_stream.len() {
            workbook_stream.splice(boundsheet_start_pos..boundsheet_actual_end, boundsheet_updated);
        }

        // Step 5: 写入 EOF 结束标志
        self.write_eof(&mut workbook_stream)?;

        // 创建CFB文件
        let file = File::create(output_path)?;
        let mut compound_file = CompoundFile::create(file)?;
        // 写入Workbook流
        let mut stream = compound_file.create_stream("Workbook")?;
        stream.write_all(&workbook_stream)?;
        stream.flush()?;
        // 关键：确保所有更改都被写入
        drop(stream);
        drop(compound_file);
        Ok(())
    }

    // --- 辅助函数 ---

    fn write_bof<W: Write>(&self, writer: &mut W, is_workbook: bool) -> Result<(), XlsError> {
        let rt = RecordType::BOF.to_u16();
        let size = 8u16; // BOF record 固定长度
        let sub_rt = if is_workbook { 0x0005 } else { 0x0010 };

        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size)?;
        writer.write_u16::<LittleEndian>(0x0600)?; // BIFF8 version
        writer.write_u16::<LittleEndian>(sub_rt)?;
        writer.write_u16::<LittleEndian>(0x0DBB)?; // Build identifier
        writer.write_u16::<LittleEndian>(0x07CC)?; // Build year

        Ok(())
    }

    fn _write_window1<W: Write>(&self, writer: &mut W) -> Result<(), XlsError> {
        let rt = RecordType::WINDOW1.to_u16();  // WINDOW1 record type
        let size = 22u16;    // WINDOW1 record fixed size

        // 窗口参数
        let x_wn = 0u16;     // 窗口左坐标
        let y_wn = 0u16;     // 窗口顶坐标
        let dx_wn = 0x10C0u16; // 窗口宽度
        let dy_wn = 0x0FF0u16; // 窗口高度
        let grbit = 0x0038u16; // 窗口标志：最大化显示
        let itab_cur = 0u16;   // 当前工作表索引
        let itab_first = 0u16; // 第一个可见工作表索引
        let c_itab_vis = 1u16; // 可见工作表数量
        let h_scroll = 0u16;   // 水平滚动位置
        let v_scroll = 0u16;   // 垂直滚动位置
        let w_tab_ratio = 0x0258u16; // 标签比例

        // 写入记录头
        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size)?;

        // 写入WINDOW1数据
        writer.write_u16::<LittleEndian>(x_wn)?;
        writer.write_u16::<LittleEndian>(y_wn)?;
        writer.write_u16::<LittleEndian>(dx_wn)?;
        writer.write_u16::<LittleEndian>(dy_wn)?;
        writer.write_u16::<LittleEndian>(grbit)?;
        writer.write_u16::<LittleEndian>(itab_cur)?;
        writer.write_u16::<LittleEndian>(itab_first)?;
        writer.write_u16::<LittleEndian>(c_itab_vis)?;
        writer.write_u16::<LittleEndian>(h_scroll)?;
        writer.write_u16::<LittleEndian>(v_scroll)?;
        writer.write_u16::<LittleEndian>(w_tab_ratio)?;

        Ok(())
    }

    fn _write_default_font<W: Write>(&self, writer: &mut W) -> Result<(), XlsError> {
        let rt = RecordType::FONT.to_u16();  // FONT record type
        let size = 22u16;    // FONT record size

        let height = 200u16;     // 字体高度 (200 twips = 10 points)
        let flags = 0u16;        // 字体标志位
        let color = 0x0000u16;   // 颜色索引 (黑色)
        let weight = 0x0190u16;  // 字体粗细 (400 = normal)
        let escapement = 0u16;   // 上下标
        let underline = 0u16;    // 下划线
        let family = 0u16;       // 字体族
        let charset = 0u16;      // 字符集
        let name_len = 5u8;      // 字体名称长度
        let name_data = b"Arial"; // 字体名称

        // 写入记录头
        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size)?;

        // 写入FONT数据
        writer.write_u16::<LittleEndian>(height)?;
        writer.write_u16::<LittleEndian>(flags)?;
        writer.write_u16::<LittleEndian>(color)?;
        writer.write_u16::<LittleEndian>(weight)?;
        writer.write_u16::<LittleEndian>(escapement)?;
        writer.write_u16::<LittleEndian>(underline)?;
        writer.write_u16::<LittleEndian>(family)?;
        writer.write_u16::<LittleEndian>(charset)?;
        writer.write_u8(name_len)?;
        writer.write_all(name_data)?;

        Ok(())
    }

    fn _write_default_xf<W: Write>(&self, writer: &mut W) -> Result<(), XlsError> {
        let rt = RecordType::XF.to_u16();  // XF record type
        let size = 30u16;    // XF record size

        let font_index = 0u16;       // 字体索引
        let format_index = 0u16;     // 格式索引
        let xtype = 0x0000u16;       // XF类型
        let align = 0x0000u16;       // 对齐方式
        let rotation = 0u16;         // 旋转角度
        let indent = 0u16;           // 缩进
        let flags = 0x0000u16;       // 标志位
        let border1 = 0x0000u32;     // 边框1
        let border2 = 0x0000u32;     // 边框2
        let back_color = 0x00000000u32; // 背景色
        let pattern_color = 0x00000000u32; // 图案颜色

        // 写入记录头
        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size)?;

        // 写入XF数据
        writer.write_u16::<LittleEndian>(font_index)?;
        writer.write_u16::<LittleEndian>(format_index)?;
        writer.write_u16::<LittleEndian>(xtype)?;
        writer.write_u16::<LittleEndian>(align)?;
        writer.write_u16::<LittleEndian>(rotation)?;
        writer.write_u16::<LittleEndian>(indent)?;
        writer.write_u16::<LittleEndian>(flags)?;
        writer.write_u32::<LittleEndian>(border1)?;
        writer.write_u32::<LittleEndian>(border2)?;
        writer.write_u32::<LittleEndian>(back_color)?;
        writer.write_u32::<LittleEndian>(pattern_color)?;

        Ok(())
    }


    // 新增方法：带偏移的 BoundSheet 写入
    fn write_boundsheet_with_offset<W: Write>(
        &self,
        writer: &mut W,
        name: &str,
        pos: u32,
    ) -> Result<(), XlsError> {
        let rt = RecordType::BOUNDSHEET.to_u16();
        let flags = 0u16; // 可视性等选项，默认隐藏状态为普通可见
        let name_grbit = 00u8;
        let name_bytes = name.as_bytes();
        let name_len = name_bytes.len() as u8;

        let size = 8 + name_len as usize;
        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size as u16)?;
        writer.write_u32::<LittleEndian>(pos)?;         // 正确偏移！
        writer.write_u16::<LittleEndian>(flags)?;
        writer.write_u8(name_len)?;
        writer.write_u8(name_grbit)?;
        writer.write_all(name_bytes)?;

        Ok(())
    }

    fn write_sst<W: Write>(&self, writer: &mut W, strings: &[String]) -> Result<(), XlsError> {
        let rt = RecordType::SST.to_u16();

        // 计算准确的总大小
        let mut string_data_size = 0usize;
        let utf16_strings: Vec<Vec<u8>> = strings.iter().map(|s| {
            let utf16_bytes: Vec<u8> = s.encode_utf16()
                .flat_map(|c| c.to_le_bytes())
                .collect();
            string_data_size += 3 + utf16_bytes.len(); // char_count(2) + flag(1) + string_data(len)
            utf16_bytes
        }).collect();

        let total_size = 8 + 4 + string_data_size; // header(8) + counts(8) + string_data

        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>((total_size - 4) as u16)?; // Size excludes first 4 bytes
        writer.write_u32::<LittleEndian>(strings.len() as u32)?; // total unique strings
        writer.write_u32::<LittleEndian>(strings.len() as u32)?; // total occurrences

        for utf16_bytes in &utf16_strings {
            let char_count = utf16_bytes.len() / 2;
            writer.write_u16::<LittleEndian>(char_count as u16)?; // character count
            writer.write_u8(0x01)?; // flag (UTF-16 little endian)
            writer.write_all(utf16_bytes)?;
        }

        Ok(())
    }

    fn collect_all_strings(&self) -> Vec<String> {
        let mut strings = std::collections::HashSet::new();
        for sheet in &self.sheets {
            for row in &sheet.rows {
                for cell in row {
                    if let Some(XlsCell::Text(s)) = cell {
                        strings.insert(s.clone());
                    }
                }
            }
        }
        strings.into_iter().collect()
    }

    fn write_sheet_data<W: Write>(
        &self,
        writer: &mut W,
        sheet: &XlsSheet,
        sst: &[String],
    ) -> Result<(), XlsError> {
        // 写入 Sheet BOF
        self.write_bof(writer, false)?;

        // 写入 DIMENSIONS record
        self.write_dimensions(writer, sheet)?;

        // 遍历每一行写入 ROW 和 CELL 数据
        for (r, row) in sheet.rows.iter().enumerate() {
            self.write_row(writer, r as u16, row.len() as u16)?;
            for (c, cell) in row.iter().enumerate() {
                if let Some(cell_data) = cell {
                    match cell_data {
                        XlsCell::Number(n) => self.write_number_cell(writer, r as u16, c as u16, *n)?,
                        XlsCell::Text(t) => {
                            let idx = sst.iter().position(|s| s == t).unwrap_or(0);
                            self.write_label_sst_cell(writer, r as u16, c as u16, idx as u32)?;
                        },
                        XlsCell::Boolean(b) => self.write_bool_cell(writer, r as u16, c as u16, *b)?,
                    }
                }
            }
        }

        // 写入 Sheet EOF
        self.write_eof(writer)?;
        Ok(())
    }

    fn write_dimensions<W: Write>(&self, writer: &mut W, sheet: &XlsSheet) -> Result<(), XlsError> {
        let rt = RecordType::DIMENSIONS.to_u16();
        let size = 14;
        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size as u16)?;

        let (max_row, max_col) = sheet.data_range().unwrap_or((0, 0));
        writer.write_u32::<LittleEndian>(0)?; // First row
        writer.write_u32::<LittleEndian>((max_row + 1) as u32)?; // Last row + 1
        writer.write_u16::<LittleEndian>(0)?; // First column
        writer.write_u16::<LittleEndian>((max_col + 1) as u16)?; // Last column + 1
        writer.write_u16::<LittleEndian>(0)?; // Reserved

        Ok(())
    }

    fn write_row<W: Write>(&self, writer: &mut W, row_index: u16, last_col: u16) -> Result<(), XlsError> {
        let rt = RecordType::Row.to_u16();
        let size = 16;
        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size as u16)?;  // 16字节
        writer.write_u16::<LittleEndian>(row_index)?;    // 0-3: row index
        writer.write_u16::<LittleEndian>(0)?;            // 4-7: first column (0)
        writer.write_u16::<LittleEndian>(last_col)?; // 8-11: last column
        writer.write_u16::<LittleEndian>(0)?;            // 12-13: row height (default)
        writer.write_u16::<LittleEndian>(0)?;            // 14-15: reserved
        writer.write_u16::<LittleEndian>(0)?;            // 14-15: reserved
        writer.write_u16::<LittleEndian>(0)?;            // 14-15: reserved
        writer.write_u16::<LittleEndian>(0)?;            // 14-15: reserved
        Ok(())
    }

    fn write_number_cell<W: Write>(&self, writer: &mut W, row: u16, col: u16, value: f64) -> Result<(), XlsError> {
        let rt = RecordType::NUMBER.to_u16();
        let size = 14;
        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size as u16)?;
        writer.write_u16::<LittleEndian>(row)?;
        writer.write_u16::<LittleEndian>(col)?;
        writer.write_u16::<LittleEndian>(0)?; // XF Index placeholder
        writer.write_f64::<LittleEndian>(value)?;
        Ok(())
    }

    fn write_label_sst_cell<W: Write>(&self, writer: &mut W, row: u16, col: u16, sst_index: u32) -> Result<(), XlsError> {
        let rt = RecordType::LABELSST.to_u16();
        let size = 10;
        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size as u16)?;
        writer.write_u16::<LittleEndian>(row)?;
        writer.write_u16::<LittleEndian>(col)?;
        writer.write_u16::<LittleEndian>(0)?; // XF Index placeholder
        writer.write_u32::<LittleEndian>(sst_index)?;
        Ok(())
    }

    fn write_bool_cell<W: Write>(&self, writer: &mut W, row: u16, col: u16, value: bool) -> Result<(), XlsError> {
        let rt = RecordType::BOOL.to_u16();
        let size = 8;
        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size as u16)?;
        writer.write_u16::<LittleEndian>(row)?;
        writer.write_u16::<LittleEndian>(col)?;
        writer.write_u16::<LittleEndian>(0)?; // XF Index placeholder
        writer.write_u8(value as u8)?;
        writer.write_u8(0)?; // padding
        Ok(())
    }

    fn write_eof<W: Write>(&self, writer: &mut W) -> Result<(), XlsError> {
        let rt = RecordType::EOF.to_u16();
        let size = 0;
        writer.write_u16::<LittleEndian>(rt)?;
        writer.write_u16::<LittleEndian>(size)?;
        Ok(())
    }
}

impl XlsWorkbook {

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

        let is_compressed = (grbit & 0x01) == 0;     // bit 0: 0=压缩, 1=未压缩
        let is_utf16 = (grbit & 0x08) == 0;          // bit 3: 0=ASCII, 1=Unicode   // 新增：检查是否压缩

        let sheet_name = if is_utf16 && !is_compressed {
            // UTF-16LE 模式：每个字符 2 字节
            let mut buf = vec![0u8; char_count * 2];
            reader.read_exact(&mut buf)?;
            let (cow, _, _) = UTF_16LE.decode(&buf);
            cow.into_owned()
        } else {
            // Latin-1 模式或压缩的UTF-16：每个字符 1 字节
            let mut buf = vec![0u8; char_count];
            reader.read_exact(&mut buf)?;
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
            for _ in 0..last_row {
                sheet.rows.push(vec![None; last_col as usize]);
            }
        }

        Ok(())
    }

    fn parse_row<R: Read + Seek>(
        reader: &mut R, current_sheet: &mut Option<XlsSheet>,
    ) -> Result<(), XlsError> {
        let row_index = reader.read_u16::<LittleEndian>()? as usize;
        let _first_col = reader.read_u16::<LittleEndian>()? as usize;
        let last_col = reader.read_u16::<LittleEndian>()? as usize;

        // 确保行存在并初始化列
        if let Some(ref mut sheet) = current_sheet {
            if sheet.rows.len() <= row_index {
                sheet.rows.resize_with(row_index, || vec![]);
            }
            sheet.rows[row_index].resize_with(last_col, || None);
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
                Ok(v) => RecordType::from_u16(v),
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
                RecordType::BOF  => Self::parse_bof(reader, &mut current_sheet, &mut sheets, &sheet_names)?,
                // BOUNDSHEET - 定义工作表的位置和名称
                RecordType::BOUNDSHEET => Self::parse_bound_sheet(reader, &mut sheet_names)?,
                // SST - 定义共享字符串
                RecordType::SST => Self::parse_sst(reader, &length, &mut sst)?,
                // DIMENSIONS - 指定工作表使用的最大行列范围
                RecordType::DIMENSIONS => Self::parse_dimensions(reader, &mut current_sheet)?,
                // ROW - 描述单行的属性（如列跨度）
                RecordType::Row => Self::parse_row(reader, &mut current_sheet)?,
                // NUMBER - 真正的 8 字节浮点数记录
                RecordType::NUMBER => Self::parse_number(reader, &mut current_sheet)?,
                // RK (0x027E / 0x007E) - XLS 最常用的数值格式（整型或压缩浮点）
                RecordType::RK | RecordType::RKOLD => Self::parse_rk(reader, &mut current_sheet)?,
                // MULRK (0x00BD) - 一行中连续的多个数字单元格
                RecordType::MULRK => Self::parse_mulrk(reader, &mut current_sheet, &length)?,
                // LABEL - 内联字符串（已弃用但仍被支持）
                RecordType::LABEL => Self::parse_label(reader, &mut current_sheet)?,
                // SST - 共享字符串
                RecordType::LABELSST => Self::parse_label_sst(reader, &mut current_sheet, &sst)?,
                // FORMULA - 公式
                RecordType::FORMULA => Self::parse_formula(reader, &mut current_sheet, &length)?,
                // EOF (End of File) - 表示当前部分结束
                RecordType::EOF => Self::parse_eof(&mut current_sheet, &mut sheets)?,
                // 忽略未知或未处理的记录
                _ => { reader.seek(SeekFrom::Current(length as i64))?; } // 跳过该记录内容
            }

            let end_pos = start_pos + length as u64;
            reader.seek(SeekFrom::Start(end_pos))?;
        }

        Ok(sheets) // 返回解析结果
    }
}

#[cfg(test)]
mod tests {
    use super::*; // 假设你已经导入了相关的结构体和方法

    /// 可视化XLS文件的记录结构
    pub fn _dump_xls_records(file_path: &str) -> Result<(), Box<dyn std::error::Error>> {
        // 打开复合文件
        let mut file = File::open(file_path)?;
        let mut cfb = CompoundFile::open(&mut file)?;

        println!("=== XLS文件 Workbook 流记录结构分析 ===");
        println!("文件路径: {}", file_path);
        println!();

        // 打开Workbook流
        let mut workbook_stream = match cfb.open_stream("Workbook") {
            Ok(stream) => stream,
            Err(_) => {
                // 尝试WorkBook（大写B）
                cfb.open_stream("WorkBook")?
            }
        };

        // 读取整个Workbook流到缓冲区
        let mut buffer = Vec::new();
        workbook_stream.read_to_end(&mut buffer)?;

        // 创建游标来遍历记录
        let mut reader = Cursor::new(buffer);
        let mut _position = 0u64;
        let _record_count = 0;

        println!("Workbook 流总大小: {} 字节", reader.get_ref().len());
        println!();

        loop {
            // 记录当前位置
            let current_pos = reader.stream_position()?;

            // 尝试读取记录头
            let record_type_code = match reader.read_u16::<LittleEndian>() {
                Ok(code) => code,
                Err(_) => {
                    println!("文件结束或读取错误");
                    break;
                }
            };

            let length = match reader.read_u16::<LittleEndian>() {
                Ok(len) => len,
                Err(_) => {
                    println!("无法读取记录长度");
                    break;
                }
            };

            let record_type = _get_record_type_name(record_type_code);

            println!("位置 {:08X}: {} (0x{:04X}) - 长度: {} 字节",
                     current_pos, record_type, record_type_code, length);

            // 显示前几个字节的数据（用于调试）
            if length > 0 && length <= 1000 {
                let mut data = vec![0u8; length as usize];
                if reader.read_exact(&mut data).is_ok() {
                    _print_hex_dump(&data, 16);
                }
            } else if length > 1000 {
                // 跳过大块数据
                reader.seek(SeekFrom::Current(length as i64))?;
                println!("  [跳过 {} 字节数据]", length);
            }

            _position = reader.stream_position()?;
            println!();
        }

        Ok(())
    }

    /// 根据记录类型码获取名称
    fn _get_record_type_name(code: u16) -> String {
        RecordType::from_u16(code).to_string()
    }

    /// 打印十六进制转储
    fn _print_hex_dump(data: &[u8], bytes_per_line: usize) {
        let mut i = 0;
        while i < data.len() {
            print!("  {:04X}: ", i);

            // 打印十六进制
            for j in 0..bytes_per_line {
                if i + j < data.len() {
                    print!("{:02X} ", data[i + j]);
                } else {
                    print!("   ");
                }
            }

            // 打印ASCII字符
            print!(" ");
            for j in 0..bytes_per_line {
                if i + j < data.len() {
                    let byte = data[i + j];
                    if byte >= 32 && byte <= 126 {
                        print!("{}", byte as char);
                    } else {
                        print!(".");
                    }
                }
            }
            println!();

            i += bytes_per_line;
        }
    }


    #[test]
    fn test_write_xls_content_analysis() -> Result<(), XlsError> {
        let input_path = "data/工作簿1.xls";       // 原始输入文件路径
        let temp_output_path = "data/工作簿1_temp.xls"; // 输出中间文件路

        // 读取原始文件
        let workbook = XlsWorkbook::new();
        let parsed_workbook = workbook.read_xls(input_path)?;

        println!("原始工作表数量: {}", parsed_workbook.sheets.len());

        // 写入文件
        parsed_workbook.write_xls(temp_output_path)?;

        // 读取写入的原始字节内容进行分析
        let written_data = std::fs::read(temp_output_path)?;
        println!("写入文件总大小: {} 字节", written_data.len());

        // 分析前几个字节（BOF记录）
        if written_data.len() >= 8 {
            println!("前8字节内容: {:02X?}", &written_data[0..8]);
            // BOF应该是 0x0809 开头
            if written_data[0] == 0x09 && written_data[1] == 0x08 {
                println!("✓ BOF 记录标识正确");
            } else {
                println!("✗ BOF 记录标识错误");
            }
        }

        // 检查是否有 BoundSheet 记录 (0x0085)
        let mut found_boundsheet = false;
        for i in 0..written_data.len().saturating_sub(2) {
            if written_data[i] == 0x85 && written_data[i+1] == 0x00 {
                println!("✓ 找到 BoundSheet 记录在位置 {}", i);
                found_boundsheet = true;
                break;
            }
        }
        if !found_boundsheet {
            println!("✗ 未找到 BoundSheet 记录");
        }

        // 检查 EOF 记录 (0x000A)
        let mut found_eof = false;
        for i in 0..written_data.len().saturating_sub(2) {
            if written_data[i] == 0x0A && written_data[i+1] == 0x00 {
                println!("✓ 找到 EOF 记录在位置 {}", i);
                found_eof = true;
                break;
            }
        }
        if !found_eof {
            println!("✗ 未找到 EOF 记录");
        }

        Ok(())
    }

    #[test]
    fn test_write_xls_stream_debug() -> Result<(), XlsError> {
        let input_path = "data/工作簿1.xls";       // 原始输入文件路径
        let temp_output_path = "data/工作簿1_temp.xls"; // 输出中间文件路

        let workbook = XlsWorkbook::new();
        let parsed_workbook = workbook.read_xls(input_path)?;

        // 写入完整文件
        parsed_workbook.write_xls(temp_output_path)?;

        let result = _dump_xls_records(input_path);
        if let Err(e) = result {
            println!("分析文件时出错: {}", e);
        }
        Ok(())
    }


    #[test]
    fn test_write_xls_roundtrip() -> Result<(), XlsError> {
        let input_path = "data/工作簿1.xls";       // 原始输入文件路径
        let temp_output_path = "data/工作簿1_temp.xls"; // 输出中间文件路径

        // 第一步：读取原始文件
        let workbook = XlsWorkbook::new();
        let parsed_workbook = workbook.read_xls(input_path)?;

        println!("原始文件工作表数量: {}", parsed_workbook.sheets.len());
        for (i, sheet) in parsed_workbook.sheets.iter().enumerate() {
            println!("  工作表 {}: {} ({} 行)", i, sheet.sheet_name, sheet.rows.len());
        }

        // 第二步：写出到临时文件
        println!("开始写入文件...");
        parsed_workbook.write_xls(temp_output_path)?;
        println!("文件写入完成");

        // 检查文件是否存在且有内容
        let metadata = std::fs::metadata(temp_output_path)?;
        println!("写入文件大小: {} 字节", metadata.len());

        // 第三步：重新读取刚刚写出的文件
        let reloaded_workbook_instance = XlsWorkbook::new();
        let reloaded_workbook = reloaded_workbook_instance.read_xls(temp_output_path)?;

        println!("重新读取的工作表数量: {}", reloaded_workbook.sheets.len());

        // 这里就会失败...
        assert_eq!(parsed_workbook.sheets.len(), reloaded_workbook.sheets.len());

        for (original_sheet, reloaded_sheet) in parsed_workbook.sheets.iter().zip(reloaded_workbook.sheets.iter()) {
            assert_eq!(original_sheet.sheet_name, reloaded_sheet.sheet_name);
            assert_eq!(original_sheet.rows.len(), reloaded_sheet.rows.len());

            for (orig_row, reload_row) in original_sheet.rows.iter().zip(reloaded_sheet.rows.iter()) {
                assert_eq!(orig_row.len(), reload_row.len());
                for (orig_cell, reload_cell) in orig_row.iter().zip(reload_row.iter()) {
                    match (orig_cell, reload_cell) {
                        (Some(XlsCell::Text(a)), Some(XlsCell::Text(b))) => assert_eq!(a, b),
                        (Some(XlsCell::Number(a)), Some(XlsCell::Number(b))) => assert_eq!(a, b),
                        (Some(XlsCell::Boolean(a)), Some(XlsCell::Boolean(b))) => assert_eq!(a, b),
                        (None, None) => {}, // Both empty cells are fine
                        _ => panic!("Mismatched cell types or values"),
                    }
                }
            }
        }

        // 清理临时文件（可选）
        std::fs::remove_file(temp_output_path).ok();

        Ok(())
    }
}
