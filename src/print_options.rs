//! Excel 打印属性配置模块
//!
//! 本模块为 `.xlsx` 导出提供页面设置/打印相关的配置结构体，
//! 并在保存时映射到 `rust_xlsxwriter` 的原生 API。
//!
//! 使用示例：
//!
//! ```rust,ignore
//! use std::collections::HashMap;
//! use xlsx_writer::print_options::{Orientation, PaperSize, PrintOptions, Scaling};
//!
//! let mut sheet_options = HashMap::new();
//! sheet_options.insert(
//!     "A".to_string(),
//!     PrintOptions {
//!         orientation: Orientation::Landscape,
//!         paper_size: PaperSize::A4,
//!         scaling: Some(Scaling::FitToPages(1, 0)),
//!         ..Default::default()
//!     },
//! );
//!
//! workbook.save_with_print_properties("out.xlsx", sheet_options)?;
//! ```

use rust_xlsxwriter::{Worksheet, XlsxError};

/// 页面方向
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum Orientation {
    /// 纵向（默认值）
    #[default]
    Portrait,
    /// 横向
    Landscape,
}

/// 纸张大小
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum PaperSize {
    /// A4（默认值）
    #[default]
    A4,
    /// A3
    A3,
    /// Letter
    Letter,
    /// Legal
    Legal,
    /// 自定义纸张尺寸，传入 `rust_xlsxwriter` 的纸张编号
    Custom(u8),
}

impl PaperSize {
    /// 转换为 `rust_xlsxwriter` 内部使用的纸张编号
    pub fn as_u8(self) -> u8 {
        match self {
            PaperSize::A4 => 9,
            PaperSize::A3 => 8,
            PaperSize::Letter => 1,
            PaperSize::Legal => 5,
            PaperSize::Custom(size) => size,
        }
    }
}

/// 打印缩放方式
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Scaling {
    /// 按百分比缩放，如 `Scale(100)` 表示 100%
    Scale(u16),
    /// 调整到 `width` × `height` 页，0 表示自动
    FitToPages(u16, u16),
}

/// 页面打印顺序
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum PageOrder {
    /// 先向下，再跨页（默认值）
    #[default]
    DownThenOver,
    /// 先跨页，再向下
    OverThenDown,
}

/// 打印区域
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct PrintArea {
    pub first_row: u32,
    pub first_col: u16,
    pub last_row: u32,
    pub last_col: u16,
}

/// 顶端标题行
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RepeatRows {
    pub first_row: u32,
    pub last_row: u32,
}

/// 左端标题列
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RepeatColumns {
    pub first_col: u16,
    pub last_col: u16,
}

/// 页边距
///
/// 未设置的字段将保持 `rust_xlsxwriter` / Excel 的默认值。
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Margins {
    pub left: Option<f64>,
    pub right: Option<f64>,
    pub top: Option<f64>,
    pub bottom: Option<f64>,
    pub header: Option<f64>,
    pub footer: Option<f64>,
}

impl Margins {
    /// 创建空的页边距配置
    pub fn new() -> Self {
        Self::default()
    }

    pub fn left(mut self, value: f64) -> Self {
        self.left = Some(value);
        self
    }

    pub fn right(mut self, value: f64) -> Self {
        self.right = Some(value);
        self
    }

    pub fn top(mut self, value: f64) -> Self {
        self.top = Some(value);
        self
    }

    pub fn bottom(mut self, value: f64) -> Self {
        self.bottom = Some(value);
        self
    }

    pub fn header(mut self, value: f64) -> Self {
        self.header = Some(value);
        self
    }

    pub fn footer(mut self, value: f64) -> Self {
        self.footer = Some(value);
        self
    }
}

/// 单个工作表的打印属性配置
#[derive(Clone, Debug, Default)]
pub struct PrintOptions {
    /// 页面方向
    pub orientation: Orientation,
    /// 纸张大小
    pub paper_size: PaperSize,
    /// 页边距
    pub margins: Margins,
    /// 打印区域
    pub print_area: Option<PrintArea>,
    /// 顶端标题行
    pub repeat_rows: Option<RepeatRows>,
    /// 左端标题列
    pub repeat_columns: Option<RepeatColumns>,
    /// 缩放方式
    pub scaling: Option<Scaling>,
    /// 水平居中
    pub center_horizontally: bool,
    /// 垂直居中
    pub center_vertically: bool,
    /// 打印网格线
    pub print_gridlines: bool,
    /// 打印行号列标
    pub print_headings: bool,
    /// 黑白打印
    pub black_and_white: bool,
    /// 草稿品质
    pub draft: bool,
    /// 起始页码
    pub first_page_number: Option<u16>,
    /// 打印顺序
    pub page_order: PageOrder,
    /// 页眉
    pub header: Option<String>,
    /// 页脚
    pub footer: Option<String>,
}

/// 将 `PrintOptions` 应用到 `rust_xlsxwriter` 的 worksheet 上
pub(crate) fn apply_print_options(
    worksheet: &mut Worksheet,
    options: &PrintOptions,
) -> Result<(), XlsxError> {
    // 方向
    match options.orientation {
        Orientation::Portrait => {
            worksheet.set_portrait();
        }
        Orientation::Landscape => {
            worksheet.set_landscape();
        }
    }

    // 纸张
    worksheet.set_paper_size(options.paper_size.as_u8());

    // 页边距（未设置传 -1.0，rust_xlsxwriter 会忽略负值）
    worksheet.set_margins(
        options.margins.left.unwrap_or(-1.0),
        options.margins.right.unwrap_or(-1.0),
        options.margins.top.unwrap_or(-1.0),
        options.margins.bottom.unwrap_or(-1.0),
        options.margins.header.unwrap_or(-1.0),
        options.margins.footer.unwrap_or(-1.0),
    );

    // 打印区域
    if let Some(area) = options.print_area {
        worksheet.set_print_area(
            area.first_row,
            area.first_col,
            area.last_row,
            area.last_col,
        )?;
    }

    // 标题行 / 标题列
    if let Some(rows) = options.repeat_rows {
        worksheet.set_repeat_rows(rows.first_row, rows.last_row)?;
    }
    if let Some(cols) = options.repeat_columns {
        worksheet.set_repeat_columns(cols.first_col, cols.last_col)?;
    }

    // 缩放
    match options.scaling {
        Some(Scaling::Scale(scale)) => {
            worksheet.set_print_scale(scale);
        }
        Some(Scaling::FitToPages(width, height)) => {
            worksheet.set_print_fit_to_pages(width, height);
        }
        None => {}
    }

    // 居中
    worksheet.set_print_center_horizontally(options.center_horizontally);
    worksheet.set_print_center_vertically(options.center_vertically);

    // 其他布尔选项
    worksheet.set_print_gridlines(options.print_gridlines);
    worksheet.set_print_headings(options.print_headings);
    worksheet.set_print_black_and_white(options.black_and_white);
    worksheet.set_print_draft(options.draft);

    // 起始页码
    if let Some(page_number) = options.first_page_number {
        worksheet.set_print_first_page_number(page_number);
    }

    // 打印顺序
    worksheet.set_page_order(matches!(options.page_order, PageOrder::DownThenOver));

    // 页眉 / 页脚
    if let Some(header) = &options.header {
        worksheet.set_header(header);
    }
    if let Some(footer) = &options.footer {
        worksheet.set_footer(footer);
    }

    Ok(())
}
