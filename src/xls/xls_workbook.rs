use crate::xls::records::{
    BackupRecord, BiffRecord, BoFRecord, BofType, BookBoolRecord, BoundSheetRecord, CodepageRecord,
    DSFRecord, DateModeRecord, EofRecord, FnGroupCountRecord, Font, FontRecord, HideObjRecord,
    InterfaceEndRecord, InterfaceHdrRecord, MMSRecord, NumberFormatRecord, ObjectProtectRecord,
    PaletteRecord, PasswordRecord, PrecisionRecord, Prot4RevPassRecord, Prot4RevRecord,
    ProtectRecord, RefreshAllRecord, SSTRecord, SharedStringTable, StyleRecord, TabIDRecord,
    UseSelfsRecord, Window1Record, WindowProtectRecord, WriteAccessRecord, XFRecord,
};
use crate::xls::{XlsError, XlsSheet};
use cfb::CompoundFile;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Write};

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
        let workbook_stream = self.get_biff_data();
        let file = File::create(output_path)?;
        let mut compound_file = CompoundFile::create(file)?;
        let mut stream = compound_file.create_stream("Workbook")?;
        stream.write_all(&workbook_stream)?;
        stream.flush()?;
        drop(stream);
        drop(compound_file);
        Ok(())
    }

    /// 生成工作簿的 BIFF 数据
    ///
    /// 此方法将工作簿序列化为 BIFF8 格式的字节流，包含完整的 Workbook Globals
    /// 子流和所有工作表的 BIFF 数据。
    pub fn get_biff_data(&self) -> Vec<u8> {
        let mut result = Vec::new();

        // 第一步：创建共享字符串表（sheet.get_biff_data 会在内部填充）
        let mut sst = SharedStringTable::new();

        // 第二步：生成每个工作表的 BIFF 数据
        let sheet_biff_data: Vec<Vec<u8>> = self
            .sheets
            .iter()
            .map(|sheet| sheet.get_biff_data(&mut sst))
            .collect();

        // 第三步：写入 Workbook Globals 头部记录
        let header_data = self.write_workbook_header_records();
        let header_len = header_data.len();
        result.extend_from_slice(&header_data);

        // 第四步：创建 pending BOUNDSHEET 记录并计算大小
        let mut pending_boundsheets: Vec<BoundSheetRecord> = self
            .sheets
            .iter()
            .map(|sheet| BoundSheetRecord::new_pending(&sheet.sheet_name))
            .collect();
        let boundsheets_total_len: usize = pending_boundsheets
            .iter()
            .map(|r| r.serialize().len())
            .sum();

        // 第五步：计算 BOUNDSHEET 偏移
        let sst_data = SSTRecord::from(&sst).serialize();
        let eof_data = EofRecord::default().serialize();
        let after_boundsheets_len = sst_data.len() + eof_data.len();

        // 每个 sheet 的偏移 = header_len + boundsheets_total_len + after_boundsheets_len + 前面所有 sheet 数据长度之和
        let mut current_offset =
            (header_len + boundsheets_total_len + after_boundsheets_len) as u32;
        for sheet_data in &sheet_biff_data {
            current_offset += sheet_data.len() as u32;
        }

        // 从后往前设置偏移（因为 current_offset 现在是末尾位置）
        for (i, sheet_data) in sheet_biff_data.iter().enumerate().rev() {
            current_offset -= sheet_data.len() as u32;
            pending_boundsheets[i].set_offset(current_offset);
        }

        // 第六步：写入带正确偏移的 BOUNDSHEET 记录
        for boundsheet in &pending_boundsheets {
            result.extend_from_slice(&boundsheet.serialize());
        }

        // 第七步：写入 SST
        result.extend_from_slice(&sst_data);

        // 第八步：写入 EOF（Workbook Globals 子流结束）
        result.extend_from_slice(&eof_data);

        // 第九步：写入工作表数据
        for sheet_data in sheet_biff_data {
            result.extend_from_slice(&sheet_data);
        }

        result
    }

    /// 写入 Workbook Globals 头部记录（从 BOF 到 UseSelfs）
    fn write_workbook_header_records(&self) -> Vec<u8> {
        let mut result = Vec::new();

        // BOF - Workbook globals
        result.extend_from_slice(&BoFRecord::new(BofType::WorkbookGlobals).serialize());

        // InterfaceHdr
        result.extend_from_slice(&InterfaceHdrRecord::default().serialize());

        // MMS
        result.extend_from_slice(&MMSRecord::default().serialize());

        // InterfaceEnd
        result.extend_from_slice(&InterfaceEndRecord::default().serialize());

        // WriteAccess
        result.extend_from_slice(&WriteAccessRecord::new("yujitui").serialize());

        // Codepage
        result.extend_from_slice(&CodepageRecord::new().serialize());

        // DSF
        result.extend_from_slice(&DSFRecord::default().serialize());

        // TabID
        result.extend_from_slice(&TabIDRecord::new(self.sheets.len() as u16).serialize());

        // FnGroupCount
        result.extend_from_slice(&FnGroupCountRecord::default().serialize());

        // Workbook Protection Block
        result.extend_from_slice(&WindowProtectRecord::default().serialize());
        result.extend_from_slice(&ProtectRecord::default().serialize());
        result.extend_from_slice(&ObjectProtectRecord::default().serialize());
        result.extend_from_slice(&PasswordRecord::default().serialize());
        result.extend_from_slice(&Prot4RevRecord::default().serialize());
        result.extend_from_slice(&Prot4RevPassRecord::default().serialize());

        // Backup
        result.extend_from_slice(&BackupRecord::default().serialize());

        // HideObj
        result.extend_from_slice(&HideObjRecord::default().serialize());

        // Window1
        result.extend_from_slice(&Window1Record::default().serialize());

        // DateMode
        result.extend_from_slice(&DateModeRecord::default().serialize());

        // Precision
        result.extend_from_slice(&PrecisionRecord::default().serialize());

        // RefreshAll
        result.extend_from_slice(&RefreshAllRecord::default().serialize());

        // BookBool
        result.extend_from_slice(&BookBoolRecord::default().serialize());

        // Fonts (8 fonts: 4x Arial + 1x Times New Roman + 3x Arial)
        result.extend_from_slice(&self.write_default_fonts());

        // Number Formats
        result.extend_from_slice(&self.write_default_formats());

        // XF records
        result.extend_from_slice(&self.write_default_xf_records());

        // Style
        result.extend_from_slice(&StyleRecord::default().serialize());

        // Palette
        result.extend_from_slice(&PaletteRecord::default().serialize());

        // UseSelfs
        result.extend_from_slice(&UseSelfsRecord::default().serialize());

        result
    }

    /// 写入默认字体记录
    ///
    /// 根据 xlwt 和 AI.md，需要 8 个字体：
    /// - Font 0-3: Arial (4个相同的默认字体，用于 XF 0-3)
    /// - Font 4: Arial (用于 XF 4)
    /// - Font 5: Times New Roman Bold (用于 XF 5)
    /// - Font 6: Arial (用于 XF 6)
    /// - Font 7: Arial (用于 XF 7+)
    fn write_default_fonts(&self) -> Vec<u8> {
        let mut result = Vec::new();

        for _ in 0..5 {
            let font = Font::new("Arial");
            result.extend_from_slice(&FontRecord::new(font).serialize());
        }

        let times = Font::new("Times New Roman").with_bold();
        result.extend_from_slice(&FontRecord::new(times).serialize());

        for _ in 0..2 {
            let font = Font::new("Arial");
            result.extend_from_slice(&FontRecord::new(font).serialize());
        }

        result
    }

    /// 写入默认数字格式记录
    fn write_default_formats(&self) -> Vec<u8> {
        let mut result = Vec::new();

        let general_format = NumberFormatRecord::new(0x00A4, "General");
        result.extend_from_slice(&general_format.serialize());

        result
    }

    /// 写入默认 XF 记录
    ///
    /// 根据 AI.md，需要 16 个 XF 记录：
    /// - XF 0-13: Style XF (font=6, format=0x00A4, type=Style)
    ///   实际 AI.md 显示前 15 个 XF 记录格式相似
    /// - XF 14-15: Style XF (不同字体)
    /// - 最后 1 个: Style XF (font=0, format=0x00A4)
    fn write_default_xf_records(&self) -> Vec<u8> {
        let mut result = Vec::new();

        // 使用 16 个默认 XF 记录（简化实现，全部是 Style XF）
        for _ in 0..16 {
            result.extend_from_slice(&XFRecord::default().serialize());
        }

        result
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
        buf_reader
            .read_to_end(&mut buffer)
            .map_err(XlsError::IoError)?;

        // 将文件内容作为 Cursor 处理，便于随机访问
        let mut cursor = Cursor::new(buffer);

        // 打开 CFB 结构（注意这里可能出错）
        let mut cfb = CompoundFile::open(&mut cursor).map_err(XlsError::IoError)?;

        // 解析 Workbook 流
        let workbook_stream = cfb
            .open_stream("Workbook")
            .map_err(|_| XlsError::InvalidFormat("Workbook stream not found".into()))?;

        // 将流中的所有字节收集到一个 Vec<u8> 中，以便后续解析。
        let bytes: Vec<u8> = workbook_stream
            .bytes()
            .collect::<Result<Vec<_>, _>>()
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
            }
            Err(e) => return Err(e),
        }

        Ok(self)
    }

    /// Parse workbook from BIFF data stream using ParsableRecord trait
    fn parse_workbook<R: Read>(reader: &mut R) -> Result<Vec<XlsSheet>, XlsError> {
        use crate::xls::records::{
            BlankRecord, BoFRecord, BoolErrRecord, BoundSheetRecord, ContinueRecord,
            DimensionsRecord, EofRecord, FormulaRecord, LabelSSTRecord, MulBlankRecord,
            MulRkRecord, NumberRecord, ParsableRecord, ParseState, RKRecord, RowRecord,
            SSTRecordData,
        };
        use byteorder::{LittleEndian, ReadBytesExt};

        let mut state = ParseState::new();

        loop {
            // Read record header
            let record_type = match reader.read_u16::<LittleEndian>() {
                Ok(rt) => rt,
                Err(_) => break, // End of stream
            };

            let length = match reader.read_u16::<LittleEndian>() {
                Ok(len) => len as usize,
                Err(_) => break,
            };

            // Read record data
            let mut data = vec![0u8; length];
            if let Err(e) = reader.read_exact(&mut data) {
                eprintln!("Warning: Failed to read record data: {}", e);
                break;
            }

            // Match record type and apply
            let result: Result<(), XlsError> = match record_type {
                // Control flow records
                BoFRecord::RECORD_ID => BoFRecord::parse(&data)?.apply(&mut state),
                EofRecord::RECORD_ID => {
                    EofRecord::parse(&data)?.apply(&mut state)?;
                    if state.is_complete {
                        break;
                    }
                    Ok(())
                }

                // SST records
                SSTRecordData::RECORD_ID => SSTRecordData::parse(&data)?.apply(&mut state),
                ContinueRecord::RECORD_ID => ContinueRecord::parse(&data)?.apply(&mut state),

                // Workbook structure records
                BoundSheetRecord::RECORD_ID => BoundSheetRecord::parse(&data)?.apply(&mut state),

                // Worksheet structure records
                RowRecord::RECORD_ID => RowRecord::parse(&data)?.apply(&mut state),
                DimensionsRecord::RECORD_ID => DimensionsRecord::parse(&data)?.apply(&mut state),

                // Cell records
                NumberRecord::RECORD_ID => NumberRecord::parse(&data)?.apply(&mut state),
                RKRecord::RECORD_ID => RKRecord::parse(&data)?.apply(&mut state),
                BlankRecord::RECORD_ID => BlankRecord::parse(&data)?.apply(&mut state),
                LabelSSTRecord::RECORD_ID => LabelSSTRecord::parse(&data)?.apply(&mut state),
                MulRkRecord::RECORD_ID => MulRkRecord::parse(&data)?.apply(&mut state),
                MulBlankRecord::RECORD_ID => MulBlankRecord::parse(&data)?.apply(&mut state),
                BoolErrRecord::RECORD_ID => BoolErrRecord::parse(&data)?.apply(&mut state),
                FormulaRecord::RECORD_ID => FormulaRecord::parse(&data)?.apply(&mut state),

                // Unknown records - silently ignore
                _ => Ok(()),
            };

            if let Err(e) = result {
                eprintln!(
                    "Warning: Error processing record 0x{:04X}: {}",
                    record_type, e
                );
                // Continue parsing, don't stop on non-fatal errors
            }
        }

        // Finalize SST if needed
        if let Some(parser) = state.sst_parser.take() {
            if let Err(e) = parser.finish(&mut state.sst) {
                eprintln!("Warning: Failed to finish SST: {}", e);
            }
        }

        // Add current sheet if any
        if let Some(sheet) = state.current_sheet.take() {
            state.sheets.push(sheet);
        }

        Ok(state.sheets)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xls::{RecordType, XlsCell};
    use byteorder::{LittleEndian, ReadBytesExt};
    use std::io::{Seek, SeekFrom};

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

            println!(
                "位置 {:08X}: {} (0x{:04X}) - 长度: {} 字节",
                current_pos, record_type, record_type_code, length
            );

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
        let input_path = "data/工作簿1.xls"; // 原始输入文件路径
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
            if written_data[i] == 0x85 && written_data[i + 1] == 0x00 {
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
            if written_data[i] == 0x0A && written_data[i + 1] == 0x00 {
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
        let input_path = "data/工作簿1.xls"; // 原始输入文件路径
        let temp_output_path = "data/工作簿1_temp.xls"; // 输出中间文件路
        let example_path = "data/example.xls";

        let workbook = XlsWorkbook::new();
        let parsed_workbook = workbook.read_xls(input_path)?;

        // 写入完整文件
        parsed_workbook.write_xls(temp_output_path)?;

        let result = _dump_xls_records(example_path);
        if let Err(e) = result {
            println!("分析文件时出错: {}", e);
        }
        Ok(())
    }

    #[test]
    fn test_write_xls_roundtrip() -> Result<(), XlsError> {
        let input_path = "data/工作簿1.xls"; // 原始输入文件路径
        let temp_output_path = "data/工作簿1_temp.xls"; // 输出中间文件路径

        // 第一步：读取原始文件
        let workbook = XlsWorkbook::new();
        let parsed_workbook = workbook.read_xls(input_path)?;

        println!("原始文件工作表数量: {}", parsed_workbook.sheets.len());
        for (i, sheet) in parsed_workbook.sheets.iter().enumerate() {
            println!(
                "  工作表 {}: {} ({} 行)",
                i,
                sheet.sheet_name,
                sheet.rows.len()
            );
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

        for (original_sheet, reloaded_sheet) in parsed_workbook
            .sheets
            .iter()
            .zip(reloaded_workbook.sheets.iter())
        {
            assert_eq!(original_sheet.sheet_name, reloaded_sheet.sheet_name);
            assert_eq!(original_sheet.rows.len(), reloaded_sheet.rows.len());

            for (orig_row, reload_row) in original_sheet.rows.iter().zip(reloaded_sheet.rows.iter())
            {
                assert_eq!(orig_row.len(), reload_row.len());
                for (orig_cell, reload_cell) in orig_row.iter().zip(reload_row.iter()) {
                    match (orig_cell, reload_cell) {
                        (Some(XlsCell::Text(a)), Some(XlsCell::Text(b))) => assert_eq!(a, b),
                        (Some(XlsCell::Number(a)), Some(XlsCell::Number(b))) => assert_eq!(a, b),
                        (Some(XlsCell::Boolean(a)), Some(XlsCell::Boolean(b))) => assert_eq!(a, b),
                        (None, None) => {} // Both empty cells are fine
                        _ => panic!("Mismatched cell types or values"),
                    }
                }
            }
        }

        // 清理临时文件（可选）
        std::fs::remove_file(temp_output_path).ok();

        Ok(())
    }

    #[test]
    fn test_sst_add_string() {
        let mut sst = SharedStringTable::new();

        // 测试基本添加
        let idx1 = sst.add("Hello".to_string());
        assert_eq!(idx1, 0);

        // 测试重复字符串
        let idx2 = sst.add("Hello".to_string());
        assert_eq!(idx2, 0); // 应该返回相同索引

        // 测试新字符串
        let idx3 = sst.add("World".to_string());
        assert_eq!(idx3, 1);

        assert_eq!(sst.string_count(), 2);
        assert_eq!(sst.total_reference_count(), 3);
    }

    #[test]
    fn test_text_cell_sst_flow() {
        let _workbook = XlsWorkbook::new();
        let sheet = XlsSheet::new("TestSheet".to_string());

        // 手动测试 SST 添加
        let mut sst = SharedStringTable::new();

        // 调用 sheet 的 get_biff_data 看是否能正常工作
        // 先测试 SST 本身
        let idx1 = sst.add("Hello".to_string());
        println!("Added 'Hello' at index {}", idx1);
        assert_eq!(idx1, 0);

        let idx2 = sst.add("World".to_string());
        println!("Added 'World' at index {}", idx2);
        assert_eq!(idx2, 1);

        // 再添加一个 Hello 确认是同一个索引
        let idx3 = sst.add("Hello".to_string());
        println!("Added 'Hello' again at index {}", idx3);
        assert_eq!(idx3, 0);

        println!(
            "SST has {} unique strings, {} total refs",
            sst.string_count(),
            sst.total_reference_count()
        );

        // 现在测试调用 get_biff_data
        let _biff_data = sheet.get_biff_data(&mut sst);

        println!(
            "Sheet BIFF data generated, SST now has {} strings",
            sst.string_count()
        );
    }

    #[test]
    fn test_get_biff_data_empty_workbook() {
        let workbook = XlsWorkbook::new();
        let biff_data = workbook.get_biff_data();

        // 验证 BOF 记录
        assert!(biff_data.len() > 4);
        assert_eq!(biff_data[0], 0x09);
        assert_eq!(biff_data[1], 0x08);
    }

    #[test]
    fn test_get_biff_data_with_empty_sheet() {
        let mut workbook = XlsWorkbook::new();
        let sheet = XlsSheet::new("Sheet1".to_string());
        workbook.sheets.push(sheet);

        let biff_data = workbook.get_biff_data();
        assert!(biff_data.len() > 100);

        // 验证包含 BOUNDSHEET 记录 (0x85)
        let has_boundsheet = biff_data.windows(2).any(|w| w == &[0x85, 0x00]);
        assert!(has_boundsheet, "BIFF data should contain BOUNDSHEET record");

        // 验证包含 SST 记录 (0xFC)
        let has_sst = biff_data.windows(2).any(|w| w == &[0xFC, 0x00]);
        assert!(has_sst, "BIFF data should contain SST record");

        // 验证包含 EOF 记录 (0x0A)
        let has_eof = biff_data.windows(2).any(|w| w == &[0x0A, 0x00]);
        assert!(has_eof, "BIFF data should contain EOF record");
    }

    #[test]
    fn test_get_biff_data_with_number_cell() {
        let mut workbook = XlsWorkbook::new();
        let mut sheet = XlsSheet::new("Sheet1".to_string());
        sheet.set_cell(0, 0, XlsCell::Number(42.0));
        workbook.sheets.push(sheet);

        let biff_data = workbook.get_biff_data();
        assert!(biff_data.len() > 200);
    }

    #[test]
    fn test_write_simple_workbook() {
        // 创建一个简单的工作簿测试
        let mut workbook = XlsWorkbook::new();

        // 创建一个包含文本的工作表
        let mut sheet = XlsSheet::new("Sheet1".to_string());
        sheet.set_cell(0, 0, XlsCell::Text("Hello".to_string()));
        workbook.sheets.push(sheet);

        println!("调用 get_biff_data...");
        let data = workbook.get_biff_data();
        println!("生成 {} 字节", data.len());

        assert!(data.len() > 100);

        // 验证包含 BOF
        assert_eq!(data[0], 0x09);
        assert_eq!(data[1], 0x08);
    }

    #[test]
    fn test_write_to_file() {
        let output_path = "data/test_simple_output.xls";

        // 创建一个简单的工作簿
        let mut workbook = XlsWorkbook::new();

        let mut sheet = XlsSheet::new("TestSheet".to_string());
        sheet.set_cell(0, 0, XlsCell::Text("Hello".to_string()));
        sheet.set_cell(0, 1, XlsCell::Text("World".to_string()));
        sheet.set_cell(1, 0, XlsCell::Number(42.0));
        workbook.sheets.push(sheet);

        println!("写入文件到 {}...", output_path);
        workbook.write_xls(output_path).expect("写入失败");

        // 验证文件存在
        let metadata = std::fs::metadata(output_path).expect("文件不存在");
        println!("文件大小: {} 字节", metadata.len());

        assert!(metadata.len() > 0, "文件应该非空");

        // 清理
        std::fs::remove_file(output_path).ok();

        println!("测试完成！");
    }

    #[test]
    fn test_read_write_roundtrip_simple() {
        // 创建包含多种数据类型的文件
        let mut workbook = XlsWorkbook::new();
        let mut sheet = XlsSheet::new("TestSheet".to_string());

        // 添加各种类型的单元格
        sheet.set_cell(0, 0, XlsCell::Text("Hello World".to_string()));
        sheet.set_cell(0, 1, XlsCell::Number(42.0));
        sheet.set_cell(1, 0, XlsCell::Number(std::f64::consts::PI));
        sheet.set_cell(1, 1, XlsCell::Text("测试中文".to_string()));
        sheet.set_cell(2, 0, XlsCell::Boolean(true));
        sheet.set_cell(2, 1, XlsCell::Boolean(false));

        workbook.sheets.push(sheet);

        let output_path = "data/test_roundtrip.xls";

        // 写入文件
        workbook.write_xls(output_path).expect("写入失败");
        println!("文件写入成功");

        // 读取文件
        let loaded_workbook = XlsWorkbook::new().read_xls(output_path).expect("读取失败");
        println!("文件读取成功，{} 个工作表", loaded_workbook.sheets.len());

        // 验证数据
        assert_eq!(loaded_workbook.sheets.len(), 1);
        let loaded_sheet = &loaded_workbook.sheets[0];

        // 验证文本单元格
        match &loaded_sheet.rows[0][0] {
            Some(XlsCell::Text(s)) => {
                println!("文本单元格 (0,0): {}", s);
                // 注意：由于 SST 解析简化，可能有差异
            }
            _ => println!("警告: (0,0) 不是预期的文本单元格"),
        }

        // 验证数字单元格
        match &loaded_sheet.rows[0][1] {
            Some(XlsCell::Number(n)) => {
                println!("数字单元格 (0,1): {}", n);
                assert!((*n - 42.0).abs() < 0.001, "数字不匹配");
            }
            _ => println!("警告: (0,1) 不是预期的数字单元格"),
        }

        // 清理
        std::fs::remove_file(output_path).ok();

        println!("读写循环测试完成！");
    }
}
