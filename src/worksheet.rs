//! Excel 工作表导出任务定义模块
//!
//! 本模块定义了 Excel 工作表导出的核心数据结构，包括：
//! - WorkSheet: 单个工作表的导出任务，由多个 SheetRegion 组成
//! - 工作表名称验证机制

use crate::cell::Cell;
use crate::error::XlsxError;
use crate::region_styles::RegionStyles;
use crate::sheet_region::SheetRegion;
use crate::xls_records::{
    row_data_to_cell_records, BiffRecord, BoFRecord, BofType, BottomMarginRecord, CalcCountRecord,
    CalcModeRecord, DBCellRecord, DefaultRowHeightRecord, DefColWidthRecord, DeltaRecord,
    DimensionsRecord, EofRecord, GridSetRecord, GutsRecord, HCenterRecord,
    IndexRecord, IterationRecord, LeftMarginRecord, PrintGridLinesRecord,
    PrintHeadersRecord, RefModeRecord, RightMarginRecord, RowRecord,
    SharedStringTable, TopMarginRecord, VCenterRecord, WSBoolRecord, Window2Record,
};
use polars::prelude::DataFrame;
use std::collections::{HashSet};

/// Excel 工作表名称中禁止出现的特殊字符集合
///
/// 根据 Microsoft Excel 规范，工作表名称不能包含以下字符：
/// - 反斜杠 `\` - 文件系统路径分隔符冲突
/// - 正斜杠 `/` - 文件系统路径分隔符冲突
/// - 问号 `?` - URL 查询参数分隔符冲突
/// - 星号 `*` - 通配符冲突
/// - 冒号 `:` - 驱动器标识符冲突
/// - 左方括号 `[` - 引用语法冲突
/// - 右方括号 `]` - 引用语法冲突
const FORBIDDEN_SHEET_NAME_CHARS: [char; 7] = ['\\', '/', '?', '*', ':', '[', ']'];

/// Excel 工作表导出任务描述符
///
/// 封装了单个工作表导出所需的所有信息。
/// 工作表由多个 SheetRegion 组成，按顺序写入。
#[derive(Debug)]
pub struct WorkSheet {
    /// 工作表显示名称
    ///
    /// 在 Excel 中显示的工作表标签名称，需符合命名规范。
    pub name: String,

    /// 区域列表
    ///
    /// 存储工作表中的所有数据区域，按添加顺序写入。
    /// 每个区域有唯一的名称用于标识。
    pub regions: Vec<SheetRegion>,
}

impl WorkSheet {
    /// 创建新的工作表
    ///
    /// # 参数
    /// * `name` - 工作表名称
    /// * `regions` - 区域列表
    ///
    /// # 返回值
    /// * `Ok(WorkSheet)` - 成功创建
    /// * `Err(XlsxError)` - 创建失败（名称无效或区域名称重复）
    pub fn new(name: impl Into<String>, regions: Vec<SheetRegion>) -> Result<Self, XlsxError> {
        let name = name.into();

        // 1. 验证工作表名称
        let trimmed = name.trim();
        if trimmed.is_empty()
            || trimmed.chars().count() > 31
            || trimmed.contains(&FORBIDDEN_SHEET_NAME_CHARS[..])
        {
            return Err(XlsxError::InvalidName(name));
        }

        // 2. 验证区域名称唯一性
        let mut name_set = HashSet::new();
        for region in &regions {
            if !name_set.insert(region.name.clone()) {
                return Err(XlsxError::GenericError(format!(
                    "Duplicate region name: {}",
                    region.name
                )));
            }
        }

        Ok(WorkSheet { name, regions })
    }

    /// 从 DataFrame 创建 WorkSheet
    ///
    /// 创建一个包含单个区域（默认名"data"）的工作表。
    ///
    /// # 参数
    /// * `df` - Polars DataFrame 数据源
    /// * `sheet_name` - 工作表显示名称
    /// * `region_name` - 区域名称（默认为"data"）
    /// * `style_map` - 可选的单元格样式映射
    /// * `merge_ranges` - 可选的合并单元格区域列表
    ///
    /// # 返回值
    /// * `Ok(WorkSheet)` - 成功创建的工作表实例
    /// * `Err(XlsxError)` - 创建过程中发生的验证错误
    pub fn from_dataframe(
        df: DataFrame,
        sheet_name: impl Into<String>,
        region_name: Option<String>,
        styles: RegionStyles,
    ) -> Result<Self, XlsxError> {
        let region_name = region_name.unwrap_or_else(|| "data".to_string());

        // 使用 SheetRegion::from_dataframe 创建区域
        let region = SheetRegion::from_dataframe(df, region_name, None, styles)?;

        // 创建工作表
        Self::new(sheet_name, vec![region])
    }

    /// 根据名称获取区域的不可变引用
    ///
    /// # 参数
    /// * `name` - 区域名称
    ///
    /// # 返回值
    /// * `Some(&SheetRegion)` - 找到的区域
    /// * `None` - 未找到
    pub fn region(&self, name: &str) -> Option<&SheetRegion> {
        self.regions.iter().find(|r| r.name == name)
    }

    /// 根据名称获取区域的可变引用
    ///
    /// # 参数
    /// * `name` - 区域名称
    ///
    /// # 返回值
    /// * `Some(&mut SheetRegion)` - 找到的区域
    /// * `None` - 未找到
    pub fn region_mut(&mut self, name: &str) -> Option<&mut SheetRegion> {
        self.regions.iter_mut().find(|r| r.name == name)
    }

    /// 添加区域到工作表
    ///
    /// # 参数
    /// * `region` - 要添加的区域
    ///
    /// # 返回值
    /// * `Ok(())` - 添加成功
    /// * `Err(XlsxError)` - 区域名称重复
    pub fn add_region(&mut self, region: SheetRegion) -> Result<(), XlsxError> {
        // 检查名称唯一性
        if self.regions.iter().any(|r| r.name == region.name) {
            return Err(XlsxError::GenericError(format!(
                "Region name '{}' already exists",
                region.name
            )));
        }

        self.regions.push(region);
        Ok(())
    }

    /// 获取所有区域的总行数
    ///
    /// 计算所有区域行数的总和。
    pub fn total_row_count(&self) -> usize {
        self.regions.iter().map(|r| r.row_count()).sum()
    }

    /// 获取所有区域的最大列数
    ///
    /// 找出所有区域中列数的最大值。
    pub fn max_col_count(&self) -> usize {
        self.regions
            .iter()
            .map(|r| r.col_count())
            .max()
            .unwrap_or(0)
    }

    /// 获取名称包含 "header" 的区域
    ///
    /// 根据区域名称过滤，返回名称中包含 "header"（不区分大小写）的区域。
    pub fn header_regions(&self) -> Vec<&SheetRegion> {
        self.regions
            .iter()
            .filter(|r| r.name.to_lowercase().contains("header"))
            .collect()
    }

    /// 获取名称包含 "footer" 的区域
    ///
    /// 根据区域名称过滤，返回名称中包含 "footer"（不区分大小写）的区域。
    pub fn footer_regions(&self) -> Vec<&SheetRegion> {
        self.regions
            .iter()
            .filter(|r| r.name.to_lowercase().contains("footer"))
            .collect()
    }

}

/// 元数据：INDEX/DBCELL 记录在 sheet BIFF 数据中的相对偏移
pub struct SheetBiffMeta {
    /// DefColWidth 记录在 sheet data 中的起始位置
    pub defcolwidth_pos: u32,
    /// INDEX 记录的 data 段起始位置（用于修补 ibXF 和 rgibRw）
    pub index_data_offset: u32,
    /// 每个 DBCELL 记录在 sheet data 中的起始位置
    pub dbcell_offsets: Vec<u32>,
}

impl WorkSheet {
    /// 生成 BIFF8 数据（用于 .xls 写入）
    ///
    /// 此方法将工作表序列化为 BIFF8 格式的字节流。
    /// 所有区域按顺序写入，坐标会自动进行偏移处理。
    ///
    /// # 参数
    /// * `sst` - 共享字符串表，用于存储文本单元格
    ///
    /// # 返回值
    /// (BIFF8 格式的字节流, INDEX/DBCELL 元数据)
    pub fn to_biff_data(&self, sst: &mut SharedStringTable) -> (Vec<u8>, SheetBiffMeta) {
        use crate::xls_records::workbook::excel_defaults::HexBiffRecord;

        // =====================================================================
        // Phase 1: 收集所有非空行
        // =====================================================================
        #[derive(Clone)]
        struct RowItem {
            abs_row: u32,
            row_data: Vec<Option<Cell>>,
        }

        let mut all_rows: Vec<RowItem> = Vec::new();
        let mut abs_row: u32 = 0;
        for region in &self.regions {
            for rel_row in &region.data {
                if rel_row.iter().any(|c| c.is_some()) {
                    all_rows.push(RowItem {
                        abs_row,
                        row_data: rel_row.clone(),
                    });
                }
                abs_row += 1;
            }
        }

        let rw_mic = all_rows.first().map(|r| r.abs_row).unwrap_or(0);
        let rw_mac = all_rows.last().map(|r| r.abs_row + 1).unwrap_or(0);

        // =====================================================================
        // Phase 2: 分组为行块（每块最多 32 行）
        // =====================================================================
        struct RowBlock {
            rows: Vec<RowItem>,
        }

        let mut blocks: Vec<RowBlock> = Vec::new();
        for row in &all_rows {
            if let Some(last) = blocks.last_mut() {
                if last.rows.len() < 32 {
                    last.rows.push(row.clone());
                    continue;
                }
            }
            let mut new_block = RowBlock { rows: Vec::new() };
            new_block.rows.push(row.clone());
            blocks.push(new_block);
        }

        // =====================================================================
        // Phase 3: 预序列化每个块的行数据，记录大小
        // =====================================================================
        struct BlockPreData {
            rows_data: Vec<(Vec<u8>, Vec<u8>)>,
            all_row_bytes: Vec<u8>,
            all_cell_bytes: Vec<u8>,
            row_sizes: Vec<u32>,
            cell_sizes: Vec<u32>,
            total_rows_size: u32,
            num_non_empty_rows: u32,
            first_cell_offsets: Vec<u16>,
            db_rtrw: u32,
        }

        let mut block_data: Vec<BlockPreData> = Vec::new();

        for block in &blocks {
            let mut bp = BlockPreData {
                rows_data: Vec::new(),
                all_row_bytes: Vec::new(),
                all_cell_bytes: Vec::new(),
                row_sizes: Vec::new(),
                cell_sizes: Vec::new(),
                total_rows_size: 0,
                num_non_empty_rows: 0,
                first_cell_offsets: Vec::new(),
                db_rtrw: 0,
            };

            let mut cumulative_cell_offset: u16 = 0;

            for item in &block.rows {
                let row_rec = RowRecord::from_row_data(item.abs_row as usize, &item.row_data);
                let row_bytes = row_rec.serialize();
                let cell_bytes =
                    row_data_to_cell_records(item.abs_row as usize, &item.row_data, 0x0F, sst);

                let row_len = row_bytes.len() as u32;
                let cell_len = cell_bytes.len() as u32;

                bp.all_row_bytes.extend_from_slice(&row_bytes);
                bp.all_cell_bytes.extend_from_slice(&cell_bytes);

                bp.row_sizes.push(row_len);
                bp.cell_sizes.push(cell_len);
                bp.total_rows_size += row_len + cell_len;
                bp.rows_data.push((row_bytes, cell_bytes));
                bp.num_non_empty_rows += 1;

                let cell_offset_start = cumulative_cell_offset;
                // 在 ROWs → CELL 布局中，rgdb 从第一组 cell 数据开始累计
                bp.first_cell_offsets.push(cell_offset_start);
                cumulative_cell_offset += cell_len as u16;
            }

            block_data.push(bp);
        }

        // =====================================================================
        // Phase 4: 计算所有记录的精确位置
        // =====================================================================
        // 记录顺序（Excel 2026 参考文件结构）：
        // [0] BOF
        // [1] INDEX ← index_data_offset
        // 设置记录：
        // [2] CalcMode (0x000D)
        // [3] CalcCount (0x000C)
        // [4] RefMode (0x000F)
        // [5] Iteration (0x0011)
        // [6] Delta (0x0010)
        // [7] SaveRecalc (0x005F)
        // [8] PrintHeaders (0x002A)
        // [9] PrintGridLines (0x002B)
        // [10] GridSet (0x0082)
        // [11] Guts (0x0080)
        // [12] DefaultRowHeight (0x0225)
        // [13] WSBool (0x0081)
        // [14] Selection (0x0014, empty)
        // [15] 0x0015 (empty)
        // [16] HCenter (0x0083)
        // [17] VCenter (0x0084)
        // [18] LeftMargin (0x0026)
        // [19] RightMargin (0x0027)
        // [20] TopMargin (0x0028)
        // [21] BottomMargin (0x0029)
        // [22] 0x004D (30B)
        // [23] SetupPage (0x00A1, 34B)
        // [24] Ext 0x089C (38B)
        // [25] DefColWidth (0x0055) ← defcolwidth_pos (ibXF)
        // [26] Dimensions (0x0200)
        // [27+] ROW (all rows first) + CELL (all cell data) + DBCELL per block
        // [Window2 (0x023E)]
        // [Ext 0x088B (16B)]
        // [Ext 0x001D (15B)]
        // [Ext 0x00EF (6B)]
        // [Ext 0x0867 (23B)]
        // [EOF]

        let bof = BoFRecord::new(BofType::Worksheet).serialize();
        let calc_mode = CalcModeRecord::default().serialize();
        let calc_count = CalcCountRecord::default().serialize();
        let ref_mode = RefModeRecord::default().serialize();
        let iteration = IterationRecord::default().serialize();
        let delta = DeltaRecord::default().serialize();
        let save_recalc = HexBiffRecord::new(0x005F, "0100").serialize();
        let print_headers = PrintHeadersRecord::default().serialize();
        let print_grid = PrintGridLinesRecord::default().serialize();
        let grid_set = GridSetRecord::default().serialize();
        let guts = GutsRecord::default().serialize();
        let def_row_height = DefaultRowHeightRecord::default().serialize();
        let wsbool = WSBoolRecord::default().serialize();
        let sel_0014 = HexBiffRecord::new(0x0014, "").serialize();
        let sel_0015 = HexBiffRecord::new(0x0015, "").serialize();
        let hcenter = HCenterRecord::default().serialize();
        let vcenter = VCenterRecord::default().serialize();
        let left_margin = LeftMarginRecord::default().serialize();
        let right_margin = RightMarginRecord::default().serialize();
        let top_margin = TopMarginRecord::default().serialize();
        let bottom_margin = BottomMarginRecord::default().serialize();
        let rec_004d = HexBiffRecord::new(0x004D, "03200100000000000000640000002c010000000000009999ffff01000000").serialize();
        let setup_page = HexBiffRecord::new(0x00A1, "0900640001000100010083002c0100002c0100009a9999999999b93f9a9999999999").serialize();
        let ext_089c = HexBiffRecord::new(0x089C, "9c0800000000000000000000000000000000000000000000000000009c990000000000000000").serialize();
        let defcolwidth = DefColWidthRecord::default().serialize();

        let total_rows = self.total_row_count();
        let max_cols = self.max_col_count();
        let dimensions = if total_rows > 0 && max_cols > 0 {
            DimensionsRecord::new(0, (total_rows - 1) as u32, 0, (max_cols - 1) as u16)
        } else {
            DimensionsRecord::default()
        };
        let dims = dimensions.serialize();

        let window2 = Window2Record::default().serialize();
        let ext_088b = HexBiffRecord::new(0x088B, "8b080000000000000000000000000200").serialize();
        let ext_001d = HexBiffRecord::new(0x001D, "030800070000000100080008000707").serialize();
        let ext_00ef = HexBiffRecord::new(0x00EF, "060037000000").serialize();
        let ext_0867 = HexBiffRecord::new(0x0867, "670800000000000000000000020001ffffffff03440000").serialize();
        let eof = EofRecord::default().serialize();

        let num_blocks = blocks.len() as u32;
        let index_record = IndexRecord::new(0, 0, 0, vec![0u32; num_blocks as usize]);
        let index_bytes = index_record.serialize();

        // =====================================================================
        // 计算累积偏移
        // =====================================================================
        let mut pos: u32 = 0;

        pos += bof.len() as u32;
        let index_data_offset = pos + 4;
        pos += index_bytes.len() as u32;

        pos += calc_mode.len() as u32;
        pos += calc_count.len() as u32;
        pos += ref_mode.len() as u32;
        pos += iteration.len() as u32;
        pos += delta.len() as u32;
        pos += save_recalc.len() as u32;
        pos += print_headers.len() as u32;
        pos += print_grid.len() as u32;
        pos += grid_set.len() as u32;
        pos += guts.len() as u32;
        pos += def_row_height.len() as u32;
        pos += wsbool.len() as u32;
        pos += sel_0014.len() as u32;
        pos += sel_0015.len() as u32;
        pos += hcenter.len() as u32;
        pos += vcenter.len() as u32;
        pos += left_margin.len() as u32;
        pos += right_margin.len() as u32;
        pos += top_margin.len() as u32;
        pos += bottom_margin.len() as u32;
        pos += rec_004d.len() as u32;
        pos += setup_page.len() as u32;
        pos += ext_089c.len() as u32;

        let defcolwidth_pos = pos;
        pos += defcolwidth.len() as u32;

        pos += dims.len() as u32;

        // 计算每个块的 DBCELL 位置和偏移
        let mut dbcell_offsets: Vec<u32> = Vec::with_capacity(num_blocks as usize);
        let ib_xf = defcolwidth_pos;
        let mut rgib_rw: Vec<u32> = Vec::with_capacity(num_blocks as usize);

        for (_i, bp) in block_data.iter_mut().enumerate() {
            let block_start = pos;

            let row_chunk: u32 = bp.row_sizes.iter().sum();
            let cell_chunk: u32 = bp.cell_sizes.iter().sum();

            // 在 ROWs → CELL 布局中:
            // block_start 指向第一行 ROW 记录
            // 所有 ROW 记录之后紧接着所有 CELL 数据
            // DBCELL 在 CELL 数据之后
            let block_end = pos + row_chunk + cell_chunk;
            let dbcell_pos = block_end;

            let db_rtrw = block_start.wrapping_sub(dbcell_pos);

            // rgdb: 相对于 block 起始位置的 CELL 数据偏移
            let mut rgdb: Vec<u16> = Vec::with_capacity(bp.rows_data.len());
            let mut cell_offset: u16 = 0;
            for j in 0..bp.rows_data.len() {
                rgdb.push(cell_offset);
                cell_offset += bp.cell_sizes[j] as u16;
            }

            bp.db_rtrw = db_rtrw;
            bp.first_cell_offsets = rgdb;

            dbcell_offsets.push(dbcell_pos);
            rgib_rw.push(0);

            let dbcell = DBCellRecord::new(db_rtrw, bp.first_cell_offsets.clone());
            let dbcell_bytes = dbcell.serialize();

            pos = dbcell_pos + dbcell_bytes.len() as u32;
        }

        // 尾部记录
        pos += window2.len() as u32;
        pos += ext_088b.len() as u32;
        pos += ext_001d.len() as u32;
        pos += ext_00ef.len() as u32;
        pos += ext_0867.len() as u32;
        pos += eof.len() as u32;

        // =====================================================================
        // Phase 5: 按顺序写入所有记录
        // =====================================================================
        let mut result = Vec::with_capacity(pos as usize);

        result.extend_from_slice(&bof);

        let index = IndexRecord::new(rw_mic, rw_mac, ib_xf, rgib_rw);
        result.extend_from_slice(&index.serialize());

        result.extend_from_slice(&calc_mode);
        result.extend_from_slice(&calc_count);
        result.extend_from_slice(&ref_mode);
        result.extend_from_slice(&iteration);
        result.extend_from_slice(&delta);
        result.extend_from_slice(&save_recalc);
        result.extend_from_slice(&print_headers);
        result.extend_from_slice(&print_grid);
        result.extend_from_slice(&grid_set);
        result.extend_from_slice(&guts);
        result.extend_from_slice(&def_row_height);
        result.extend_from_slice(&wsbool);
        result.extend_from_slice(&sel_0014);
        result.extend_from_slice(&sel_0015);
        result.extend_from_slice(&hcenter);
        result.extend_from_slice(&vcenter);
        result.extend_from_slice(&left_margin);
        result.extend_from_slice(&right_margin);
        result.extend_from_slice(&top_margin);
        result.extend_from_slice(&bottom_margin);
        result.extend_from_slice(&rec_004d);
        result.extend_from_slice(&setup_page);
        result.extend_from_slice(&ext_089c);
        result.extend_from_slice(&defcolwidth);
        result.extend_from_slice(&dims);

        // 按 Row Block 写入：每个块内 ROW → CELL → DBCELL
        for bp in &block_data {
            result.extend_from_slice(&bp.all_row_bytes);
            result.extend_from_slice(&bp.all_cell_bytes);
            let dbcell = DBCellRecord::new(bp.db_rtrw, bp.first_cell_offsets.clone());
            result.extend_from_slice(&dbcell.serialize());
        }

        result.extend_from_slice(&window2);
        result.extend_from_slice(&ext_088b);
        result.extend_from_slice(&ext_001d);
        result.extend_from_slice(&ext_00ef);
        result.extend_from_slice(&ext_0867);
        result.extend_from_slice(&eof);

        let meta = SheetBiffMeta {
            defcolwidth_pos,
            index_data_offset,
            dbcell_offsets,
        };

        (result, meta)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::Cell;
    use polars::prelude::NamedFrom;
    use polars::prelude::Series;

    #[test]
    fn test_worksheet_new() {
        let region = SheetRegion::new("data", vec![vec![Some(Cell::Text("Test".to_string()))]]);
        let sheet = WorkSheet::new("TestSheet", vec![region]).unwrap();

        assert_eq!(sheet.name, "TestSheet");
        assert_eq!(sheet.regions.len(), 1);
    }

    #[test]
    fn test_worksheet_name_validation() {
        // 空名称
        let region = SheetRegion::new("data", vec![vec![Some(Cell::Text("Test".to_string()))]]);
        let result = WorkSheet::new("", vec![region.clone()]);
        assert!(result.is_err());

        // 超过31字符
        let long_name = "a".repeat(32);
        let result = WorkSheet::new(long_name, vec![region.clone()]);
        assert!(result.is_err());

        // 包含非法字符
        let result = WorkSheet::new("Test/Sheet", vec![region.clone()]);
        assert!(result.is_err());
    }

    #[test]
    fn test_worksheet_duplicate_region_name() {
        let region1 = SheetRegion::new("data", vec![vec![Some(Cell::Text("Test1".to_string()))]]);
        let region2 = SheetRegion::new("data", vec![vec![Some(Cell::Text("Test2".to_string()))]]);
        let result = WorkSheet::new("TestSheet", vec![region1, region2]);
        assert!(result.is_err());
    }

    #[test]
    fn test_worksheet_from_dataframe() {
        let columns: Vec<polars::frame::column::Column> = vec![
            Series::new("Name".into(), vec!["Alice", "Bob"]).into(),
            Series::new("Age".into(), vec![30_i64, 25]).into(),
        ];
        let df = DataFrame::new(2, columns).unwrap();

        let sheet = WorkSheet::from_dataframe(df, "Test", None, RegionStyles::new()).unwrap();

        assert_eq!(sheet.name, "Test");
        assert_eq!(sheet.regions.len(), 1);
        assert_eq!(sheet.regions[0].name, "data");
        assert_eq!(sheet.total_row_count(), 3); // 1 header + 2 data rows
    }

    #[test]
    fn test_worksheet_region_access() {
        let region1 =
            SheetRegion::new("header", vec![vec![Some(Cell::Text("Header".to_string()))]]);
        let region2 = SheetRegion::new("data", vec![vec![Some(Cell::Text("Data".to_string()))]]);
        let sheet = WorkSheet::new("TestSheet", vec![region1, region2]).unwrap();

        // 测试 region()
        assert!(sheet.region("header").is_some());
        assert!(sheet.region("data").is_some());
        assert!(sheet.region("nonexistent").is_none());

        // 测试 region_mut()
        let mut sheet_mut = sheet;
        assert!(sheet_mut.region_mut("header").is_some());
    }

    #[test]
    fn test_worksheet_add_region() {
        let region1 = SheetRegion::new("data1", vec![vec![Some(Cell::Text("Test1".to_string()))]]);
        let mut sheet = WorkSheet::new("TestSheet", vec![region1]).unwrap();

        let region2 = SheetRegion::new("data2", vec![vec![Some(Cell::Text("Test2".to_string()))]]);
        assert!(sheet.add_region(region2).is_ok());
        assert_eq!(sheet.regions.len(), 2);

        // 添加重复名称应该失败
        let region3 = SheetRegion::new("data1", vec![vec![Some(Cell::Text("Test3".to_string()))]]);
        assert!(sheet.add_region(region3).is_err());
    }

    #[test]
    fn test_worksheet_total_row_count() {
        let region1 = SheetRegion::new(
            "r1",
            vec![
                vec![Some(Cell::Text("H1".to_string()))],
                vec![Some(Cell::Text("D1".to_string()))],
            ],
        );
        let region2 = SheetRegion::new(
            "r2",
            vec![
                vec![Some(Cell::Text("H2".to_string()))],
                vec![Some(Cell::Text("D2".to_string()))],
                vec![Some(Cell::Text("D3".to_string()))],
            ],
        );
        let sheet = WorkSheet::new("TestSheet", vec![region1, region2]).unwrap();

        assert_eq!(sheet.total_row_count(), 5);
    }

    #[test]
    fn test_worksheet_max_col_count() {
        let region1 = SheetRegion::new(
            "r1",
            vec![vec![
                Some(Cell::Text("A".to_string())),
                Some(Cell::Text("B".to_string())),
            ]],
        );
        let region2 = SheetRegion::new(
            "r2",
            vec![vec![
                Some(Cell::Text("A".to_string())),
                Some(Cell::Text("B".to_string())),
                Some(Cell::Text("C".to_string())),
            ]],
        );
        let sheet = WorkSheet::new("TestSheet", vec![region1, region2]).unwrap();

        assert_eq!(sheet.max_col_count(), 3);
    }

    #[test]
    fn test_worksheet_header_footer_regions() {
        let header_region = SheetRegion::new(
            "my_header",
            vec![vec![Some(Cell::Text("Header".to_string()))]],
        );
        let data_region =
            SheetRegion::new("data", vec![vec![Some(Cell::Text("Data".to_string()))]]);
        let footer_region = SheetRegion::new(
            "my_footer",
            vec![vec![Some(Cell::Text("Footer".to_string()))]],
        );
        let sheet =
            WorkSheet::new("TestSheet", vec![header_region, data_region, footer_region]).unwrap();

        let headers = sheet.header_regions();
        assert_eq!(headers.len(), 1);
        assert_eq!(headers[0].name, "my_header");

        let footers = sheet.footer_regions();
        assert_eq!(footers.len(), 1);
        assert_eq!(footers[0].name, "my_footer");
    }
}
