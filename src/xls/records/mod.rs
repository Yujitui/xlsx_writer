//! BIFF记录模块
//!
//! 实现Excel BIFF8格式的所有记录类型
//!
//! ## 模块组织
//!
//! - `biff_record`: BiffRecord trait定义
//! - `bof_record` / `eof_record`: 文件开始/结束记录
//! - `utils`: 字符串编码工具函数
//! - `workbook`: 工作簿级记录（SST、XF、Font等）
//! - `worksheet`: 工作表级记录（单元格、行、窗格等）

pub mod workbook;
pub mod worksheet;

pub mod biff_record;
pub mod bof_record;
pub mod eof_record;
pub mod parseable;
pub mod utils;

pub use utils::{encode_biff_string_v1, encode_biff_string_v2};

pub use biff_record::BiffRecord;
pub use bof_record::{BoFRecord, BofType};
pub use eof_record::EofRecord;
pub use parseable::{decode_rk_value, ParsableRecord, ParseState, SSTParserState};

pub use workbook::{
    backup_record::BackupRecord,
    book_bool_record::BookBoolRecord,
    boundsheet_record::{BoundSheetRecord, SheetType, SheetVisibility},
    codepage_record::CodepageRecord,
    country_record::CountryRecord,
    date_mode_record::DateModeRecord,
    dsf_record::DSFRecord,
    fn_group_count_record::FnGroupCountRecord,
    font_record::{Font, FontRecord},
    hide_obj_record::HideObjRecord,
    interface_end_record::InterfaceEndRecord,
    interface_hdr_record::InterfaceHdrRecord,
    mms_record::MMSRecord,
    number_format_record::NumberFormatRecord,
    object_protect_record::ObjectProtectRecord,
    palette_record::PaletteRecord,
    password_record::PasswordRecord,
    precision_record::PrecisionRecord,
    prot_4_rev_pass_record::Prot4RevPassRecord,
    prot_4_rev_record::Prot4RevRecord,
    protect_record::ProtectRecord,
    refresh_all_record::RefreshAllRecord,
    sst_record::{ContinueRecord, SSTRecord, SSTRecordData, SharedStringTable},
    style_record::StyleRecord,
    tab_id_record::TabIDRecord,
    use_selfs_record::UseSelfsRecord,
    window1_record::Window1Record,
    window_protect_record::WindowProtectRecord,
    write_access_record::WriteAccessRecord,
    xf_record::{Alignment, Borders, Pattern, Protection, XFRecord, XFType, XF},
};

pub use worksheet::{
    // 其他记录
    bottom_margin_record::BottomMarginRecord,
    calc_count_record::CalcCountRecord,
    calc_mode_record::CalcModeRecord,
    cell_records::{
        row_data_to_cell_records, BlankRecord, BoolErrRecord, FormulaRecord, LabelSSTRecord,
        MulBlankRecord, MulRkRecord, NumberRecord, RKRecord,
    },
    default_row_height_record::DefaultRowHeightRecord,
    delta_record::DeltaRecord,
    dimensions_record::DimensionsRecord,
    footer_record::FooterRecord,
    grid_set_record::GridSetRecord,
    guts_record::GutsRecord,
    hcenter_record::HCenterRecord,
    header_record::HeaderRecord,
    horizontal_page_breaks_record::{HorizontalPageBreaksRecord, PageBreak},
    iteration_record::IterationRecord,
    left_margin_record::LeftMarginRecord,
    merged_cells_record::{CellRange, MergedCellsRecord},
    // 重名记录使用别名
    object_protect_record::ObjectProtectRecord as WorksheetObjectProtectRecord,
    panes_record::PanesRecord,
    password_record::{password_hash, PasswordRecord as WorksheetPasswordRecord},
    print_grid_lines_record::PrintGridLinesRecord,
    print_headers_record::PrintHeadersRecord,
    protect_record::ProtectRecord as WorksheetProtectRecord,
    ref_mode_record::RefModeRecord,
    right_margin_record::RightMarginRecord,
    row_record::RowRecord,
    save_recalc_record::SaveRecalcRecord,
    scen_protect_record::ScenProtectRecord,
    setup_page_record::SetupPageRecord,
    top_margin_record::TopMarginRecord,
    vcenter_record::VCenterRecord,
    vertical_page_breaks_record::{VerticalPageBreak, VerticalPageBreaksRecord},
    window2_record::Window2Record,
    window_protect_record::WindowProtectRecord as WorksheetWindowProtectRecord,
    wsbool_record::WSBoolRecord,
};
