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
    DimensionsRecord, EofRecord, FooterRecord, GridSetRecord, GutsRecord, HCenterRecord,
    HeaderRecord, IndexRecord, IterationRecord, LeftMarginRecord, PrintGridLinesRecord,
    PrintHeadersRecord, RefModeRecord, RightMarginRecord, RowRecord, ScenProtectRecord,
    SetupPageRecord, SharedStringTable, TopMarginRecord, VCenterRecord, WSBoolRecord, Window2Record,
    WorksheetObjectProtectRecord, WorksheetProtectRecord, WorksheetWindowProtectRecord,
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
            rows_data: Vec<(Vec<u8>, Vec<u8>)>, // (RowRecord bytes, CellRecord bytes)
            row_sizes: Vec<u32>,                 // total size of ROW + CELL for each row
            total_rows_size: u32,                // sum of all row_sizes
            num_non_empty_rows: u32,             // rows with cell data in this block
            first_cell_offsets: Vec<u16>,        // rgdb entries
            db_rtrw: u32,                        // offset from DBCELL to first ROW
        }

        let mut block_data: Vec<BlockPreData> = Vec::new();

        for block in &blocks {
            let mut bp = BlockPreData {
                rows_data: Vec::new(),
                row_sizes: Vec::new(),
                total_rows_size: 0,
                num_non_empty_rows: 0,
                first_cell_offsets: Vec::new(),
                db_rtrw: 0,
            };

            for item in &block.rows {
                let row_rec = RowRecord::from_row_data(item.abs_row as usize, &item.row_data);
                let row_bytes = row_rec.serialize();
                let cell_bytes =
                    row_data_to_cell_records(item.abs_row as usize, &item.row_data, 0, sst);
                let row_total = (row_bytes.len() + cell_bytes.len()) as u32;

                bp.total_rows_size += row_total;
                bp.row_sizes.push(row_total);
                bp.rows_data.push((row_bytes, cell_bytes));
                bp.num_non_empty_rows += 1;
            }

            block_data.push(bp);
        }

        // =====================================================================
        // Phase 4: 计算所有记录的精确位置
        // =====================================================================

        // 固定大小记录：BOF + 6 个设置记录 + Guts + DefaultRowHeight + WSBool
        // + DefColWidth + INDEX + Dimensions + 4 个打印记录 + 7 个边距记录
        // + SetupPage + 4 个保护记录 + Window2 + EOF
        // 这些大部分使用默认值，我们实际计算以确保准确

        let bof = BoFRecord::new(BofType::Worksheet).serialize();
        let calc_mode = CalcModeRecord::default().serialize();
        let calc_count = CalcCountRecord::default().serialize();
        let ref_mode = RefModeRecord::default().serialize();
        let delta = DeltaRecord::default().serialize();
        let iteration = IterationRecord::default().serialize();
        let guts = GutsRecord::default().serialize();
        let def_row_height = DefaultRowHeightRecord::default().serialize();
        let wsbool = WSBoolRecord::default().serialize();
        let defcolwidth = DefColWidthRecord::default().serialize();

        let total_rows = self.total_row_count();
        let max_cols = self.max_col_count();
        let dimensions = if total_rows > 0 && max_cols > 0 {
            DimensionsRecord::new(0, (total_rows - 1) as u32, 0, (max_cols - 1) as u16)
        } else {
            DimensionsRecord::default()
        };
        let dims = dimensions.serialize();

        let print_headers = PrintHeadersRecord::default().serialize();
        let print_grid = PrintGridLinesRecord::default().serialize();
        let grid_set = GridSetRecord::default().serialize();
        let header = HeaderRecord::default().serialize();
        let footer = FooterRecord::default().serialize();
        let hcenter = HCenterRecord::default().serialize();
        let vcenter = VCenterRecord::default().serialize();
        let left_margin = LeftMarginRecord::default().serialize();
        let right_margin = RightMarginRecord::default().serialize();
        let top_margin = TopMarginRecord::default().serialize();
        let bottom_margin = BottomMarginRecord::default().serialize();
        let setup_page = SetupPageRecord::default().serialize();
        let ws_protect = WorksheetProtectRecord::default().serialize();
        let ws_window_protect = WorksheetWindowProtectRecord::default().serialize();
        let scen_protect = ScenProtectRecord::default().serialize();
        let obj_protect = WorksheetObjectProtectRecord::default().serialize();

        let window2 = Window2Record::default().serialize();
        let eof = EofRecord::default().serialize();

        // 计算 INDEX 记录大小：reserved(4) + rwMic(4) + rwMac(4) + ibXF(4) + N*4
        let num_blocks = blocks.len() as u32;
        let index_record = IndexRecord::new(0, 0, 0, vec![0u32; num_blocks as usize]);
        let index_bytes = index_record.serialize();

        // =====================================================================
        // 计算累积偏移（单位：字节）
        // =====================================================================
        // 以下是 sheet 内各记录的相对位置：
        // [BOF]
        // [INDEX] ← index_data_offset (INDEX 的 data 段)
        // [CalcMode]  [CalcCount]  [RefMode]  [Delta]  [Iteration]
        // [Guts]  [DefaultRowHeight]  [WSBool]
        // [DefColWidth] ← defcolwidth_pos
        // [DIMENSIONS]
        // [Print/Protection records]
        // [RowBlock1: ROW + CELL records ...]
        // [DBCELL1]
        // [RowBlock2: ROW + CELL records ...]
        // [DBCELL2]
        // ...
        // [Window2]
        // [EOF]

        let mut pos: u32 = 0;

        // BOF
        pos += bof.len() as u32;

        // INDEX 紧接着 BOF（4 字节头部之后）
        let index_data_offset = pos + 4; // skip INDEX header (id + len)
        pos += index_bytes.len() as u32;

        // 设置记录
        pos += calc_mode.len() as u32;
        pos += calc_count.len() as u32;
        pos += ref_mode.len() as u32;
        pos += delta.len() as u32;
        pos += iteration.len() as u32;
        pos += guts.len() as u32;
        pos += def_row_height.len() as u32;
        pos += wsbool.len() as u32;

        // DefColWidth
        let defcolwidth_pos = pos;
        pos += defcolwidth.len() as u32;

        // Dimensions + 打印记录 + 保护记录
        pos += dims.len() as u32;
        pos += print_headers.len() as u32;
        pos += print_grid.len() as u32;
        pos += grid_set.len() as u32;
        pos += header.len() as u32;
        pos += footer.len() as u32;
        pos += hcenter.len() as u32;
        pos += vcenter.len() as u32;
        pos += left_margin.len() as u32;
        pos += right_margin.len() as u32;
        pos += top_margin.len() as u32;
        pos += bottom_margin.len() as u32;
        pos += setup_page.len() as u32;
        pos += ws_protect.len() as u32;
        pos += ws_window_protect.len() as u32;
        pos += scen_protect.len() as u32;
        pos += obj_protect.len() as u32;

        // 计算每个块的 DBCELL 位置和偏移
        let mut dbcell_offsets: Vec<u32> = Vec::with_capacity(num_blocks as usize);
        let ib_xf = defcolwidth_pos;
        let mut rgib_rw: Vec<u32> = Vec::with_capacity(num_blocks as usize);

        for (_i, bp) in block_data.iter_mut().enumerate() {
            // 块起始位置 (第一个 ROW 记录的位置)
            let block_start = pos;

            // 计算 first_cell_offsets (rgdb)
            // 布局: [Row_0][Cell_0][Row_1][Cell_1]...
            // rgdb[0] = 0 (first CELL immediately follows first ROW)
            // rgdb[i] = cell_bytes[i-1].len() + row_bytes[i].len() (for i >= 1)
            let mut rgdb: Vec<u16> = Vec::with_capacity(bp.rows_data.len());
            for j in 0..bp.rows_data.len() {
                if j == 0 {
                    rgdb.push(0);
                } else {
                    let prev_cell_len = bp.rows_data[j - 1].1.len() as u16;
                    let cur_row_len = bp.rows_data[j].0.len() as u16;
                    rgdb.push(prev_cell_len + cur_row_len);
                }
            }

            // DBCELL 记录的位置
            let block_end = pos + bp.total_rows_size;
            let dbcell_pos = block_end;

            // dbRtrw: 从 DBCELL 起始位置到第一个 ROW 记录的偏移
            // (负偏移 → u32 回绕)
            let db_rtrw = block_start.wrapping_sub(dbcell_pos);

            bp.db_rtrw = db_rtrw;
            bp.first_cell_offsets = rgdb;

            dbcell_offsets.push(dbcell_pos);
            rgib_rw.push(0); // 占位，稍后在 workbook 中修补

            // 准备 DBCELL 序列化
            let dbcell = DBCellRecord::new(db_rtrw, bp.first_cell_offsets.clone());
            let dbcell_bytes = dbcell.serialize();

            pos = dbcell_pos + dbcell_bytes.len() as u32;
        }

        // 最后加上 Window2 + EOF
        pos += window2.len() as u32;
        pos += eof.len() as u32;

        // =====================================================================
        // Phase 5: 按顺序写入所有记录
        // =====================================================================
        let mut result = Vec::with_capacity(pos as usize);

        result.extend_from_slice(&bof);

        // INDEX 紧接着 BOF (with relative offsets; workbook will patch absolute FilePointers)
        let index = IndexRecord::new(rw_mic, rw_mac, ib_xf, rgib_rw);
        result.extend_from_slice(&index.serialize());

        result.extend_from_slice(&calc_mode);
        result.extend_from_slice(&calc_count);
        result.extend_from_slice(&ref_mode);
        result.extend_from_slice(&delta);
        result.extend_from_slice(&iteration);
        result.extend_from_slice(&guts);
        result.extend_from_slice(&def_row_height);
        result.extend_from_slice(&wsbool);

        // DefColWidth
        result.extend_from_slice(&defcolwidth);

        // Dimensions
        result.extend_from_slice(&dims);

        // 打印和保护记录
        result.extend_from_slice(&print_headers);
        result.extend_from_slice(&print_grid);
        result.extend_from_slice(&grid_set);
        result.extend_from_slice(&header);
        result.extend_from_slice(&footer);
        result.extend_from_slice(&hcenter);
        result.extend_from_slice(&vcenter);
        result.extend_from_slice(&left_margin);
        result.extend_from_slice(&right_margin);
        result.extend_from_slice(&top_margin);
        result.extend_from_slice(&bottom_margin);
        result.extend_from_slice(&setup_page);
        result.extend_from_slice(&ws_protect);
        result.extend_from_slice(&ws_window_protect);
        result.extend_from_slice(&scen_protect);
        result.extend_from_slice(&obj_protect);

        // 行块数据 + DBCELL
        for bp in &block_data {
            for (ref row_bytes, ref cell_bytes) in &bp.rows_data {
                result.extend_from_slice(row_bytes);
                result.extend_from_slice(cell_bytes);
            }
            let dbcell = DBCellRecord::new(bp.db_rtrw, bp.first_cell_offsets.clone());
            result.extend_from_slice(&dbcell.serialize());
        }

        result.extend_from_slice(&window2);
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
