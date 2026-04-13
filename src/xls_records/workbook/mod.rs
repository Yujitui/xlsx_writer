//! 工作簿记录模块
//!
//! 包含BIFF格式中工作簿(Workbook)相关的所有记录类型

pub mod backup_record;
pub mod book_bool_record;
pub mod boundsheet_record;
pub mod codepage_record;
pub mod country_record;
pub mod date_mode_record;
pub mod dsf_record;
pub mod fn_group_count_record;
pub mod font_record;
pub mod hide_obj_record;
pub mod interface_end_record;
pub mod interface_hdr_record;
pub mod mms_record;
pub mod number_format_record;
pub mod object_protect_record;
pub mod palette_record;
pub mod password_record;
pub mod precision_record;
pub mod prot_4_rev_pass_record;
pub mod prot_4_rev_record;
pub mod protect_record;
pub mod refresh_all_record;
pub mod sst_record;
pub mod style_record;
pub mod tab_id_record;
pub mod use_selfs_record;
pub mod window1_record;
pub mod window_protect_record;
pub mod write_access_record;
pub mod xf_record;

pub use crate::xls_records::BiffRecord;
pub use crate::xls_records::{encode_biff_string_v1, encode_biff_string_v2};
