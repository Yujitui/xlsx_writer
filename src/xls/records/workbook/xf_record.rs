use super::BiffRecord;

/// 对齐方式结构体
///
/// ## 作用
///
/// 定义单元格文本的对齐方式，包括水平对齐、垂直对齐、自动换行等。
///
/// ## 参数说明
///
/// - `horz`: 水平对齐方式
///   - 0 = 常规, 1 = 左对齐, 2 = 居中, 3 = 右对齐
///   - 4 = 填充, 5 = 两端对齐, 6 = 跨列居中, 7 = 分散对齐
/// - `vert`: 垂直对齐方式
///   - 0 = 靠上, 1 = 居中, 2 = 靠下, 3 = 两端对齐, 4 = 分散对齐
/// - `wrap`: 自动换行（0=不换行, 1=换行）
/// - `rota`: 旋转角度（0-90度, 255=堆叠文字）
/// - `inde`: 缩进级别（0-15）
/// - `shri`: 缩小填充（0=否, 1=是）
/// - `merg`: 合并单元格（0=否, 1=是）
/// - `dire`: 文字方向（0=根据上下文, 1=从左到右, 2=从右到左）
#[derive(Debug, Clone)]
pub struct Alignment {
    pub horz: u8, // 0=general, 1=left, 2=center, 3=right, 4=fill, 5=justify, 6=center_across, 7=distributed
    pub vert: u8, // 0=top, 1=center, 2=bottom, 3=justify, 4=distributed
    pub wrap: u8, // 0=no wrap, 1=wrap
    pub rota: u8, // rotation angle (0-90, 255=stacked)
    pub inde: u8, // indent level (0-15)
    pub shri: u8, // shrink to fit (0=no, 1=yes)
    pub merg: u8, // merge cells (0=no, 1=yes)
    pub dire: u8, // text direction (0=context, 1=left-to-right, 2=right-to-left)
}

impl Default for Alignment {
    fn default() -> Self {
        Alignment {
            horz: 0,
            vert: 2,
            wrap: 0,
            rota: 0,
            inde: 0,
            shri: 0,
            merg: 0,
            dire: 0,
        }
    }
}

/// 边框结构体
///
/// ## 作用
///
/// 定义单元格的边框样式，包括边框线类型和颜色。
///
/// ## 参数说明
///
/// - `left`, `right`, `top`, `bottom`, `diag`: 边框线类型
///   - 0 = 无, 1 = 细, 2 = 中, 3 = 虚线, 4 = 点虚线, 5 = 粗
///   - 6 = 双线, 7 = hair, 8-13 = 特定样式
/// - `left_colour`, `right_colour`, `top_colour`, `bottom_colour`, `diag_colour`: 边框颜色索引
/// - `need_diag1`, `need_diag2`: 是否显示对角线（0=否, 1=是）
#[derive(Debug, Clone)]
pub struct Borders {
    pub left: u8,
    pub right: u8,
    pub top: u8,
    pub bottom: u8,
    pub diag: u8,
    pub left_colour: u8,
    pub right_colour: u8,
    pub top_colour: u8,
    pub bottom_colour: u8,
    pub diag_colour: u8,
    pub need_diag1: u8,
    pub need_diag2: u8,
}

impl Default for Borders {
    fn default() -> Self {
        Borders {
            left: 0,
            right: 0,
            top: 0,
            bottom: 0,
            diag: 0,
            left_colour: 0,
            right_colour: 0,
            top_colour: 0,
            bottom_colour: 0,
            diag_colour: 0,
            need_diag1: 0,
            need_diag2: 0,
        }
    }
}

/// 填充图案结构体
///
/// ## 作用
///
/// 定义单元格的背景填充图案和颜色。
///
/// ## 参数说明
///
/// - `pattern`: 图案样式（0=无, 1=实色, 2=50%灰色, 3=75%灰色, 4=25%灰色,
///   5=横条纹, 6=竖条纹, 7=反向对角条纹, 8=对角条纹, 9= thick diagonal crosshatch,
///   10=thin horizontal stripe, 11=thin vertical stripe, 12=thin reverse diagonal stripe,
///   13=thin diagonal stripe, 14=thin diagonal crosshatch, 15=thick horizontal crosshatch,
///   16=thick vertical crosshatch, 17=thick reverse diagonal crosshatch, 18=thick diagonal crosshatch）
/// - `pattern_fore_colour`: 前景色索引
/// - `pattern_back_colour`: 背景色索引
#[derive(Debug, Clone)]
pub struct Pattern {
    pub pattern: u8,
    pub pattern_fore_colour: u8,
    pub pattern_back_colour: u8,
}

impl Default for Pattern {
    fn default() -> Self {
        Pattern {
            pattern: 0,
            pattern_fore_colour: 0,
            pattern_back_colour: 0,
        }
    }
}

/// 保护结构体
///
/// ## 作用
///
/// 定义单元格的锁定和公式隐藏状态。
///
/// ## 参数说明
///
/// - `cell_locked`: 单元格锁定（0=未锁定, 1=锁定）
/// - `formula_hidden`: 公式隐藏（0=显示公式, 1=隐藏公式）
#[derive(Debug, Clone)]
pub struct Protection {
    pub cell_locked: u8,
    pub formula_hidden: u8,
}

impl Default for Protection {
    fn default() -> Self {
        Protection {
            cell_locked: 1,
            formula_hidden: 0,
        }
    }
}

/// XF (扩展格式) 结构体
///
/// ## 作用
///
/// XF结构体是Excel单元格格式的核心，包含了单元格的所有样式信息：
/// 字体、边框、填充图案、保护设置等。每个XF通过索引被单元格引用。
///
/// ## 参数说明
///
/// - `font_idx`: 字体索引（指向FontRecord）
/// - `format_idx`: 数字格式索引（指向NumberFormatRecord，164及以下为内置格式）
/// - `alignment`: 对齐方式
/// - `borders`: 边框样式
/// - `pattern`: 填充图案
/// - `protection`: 保护设置
#[derive(Debug, Clone)]
pub struct XF {
    pub font_idx: u16,
    pub format_idx: u16,
    pub alignment: Alignment,
    pub borders: Borders,
    pub pattern: Pattern,
    pub protection: Protection,
}

impl Default for XF {
    fn default() -> Self {
        XF {
            font_idx: 0,
            format_idx: 164,
            alignment: Alignment::default(),
            borders: Borders::default(),
            pattern: Pattern::default(),
            protection: Protection::default(),
        }
    }
}

/// XF类型枚举
///
/// ## 作用
///
/// 区分单元格XF和样式XF。
///
/// ## 参数说明
///
/// - `Cell`: 单元格格式XF
/// - `Style`: 样式XF（用于内置样式定义）
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XFType {
    Cell,
    Style,
}

/// XFRecord 记录
///
/// 作用：存储扩展格式（XF）定义
///
/// XFRecord是Excel BIFF格式中的扩展格式记录（ID: 0x00E0），用于存储单元格或样式的格式信息。
/// 每个XFRecord包含一个XF结构体，定义了完整的单元格格式属性。
///
/// ## 参数说明
///
/// - `xf`: XF结构体，包含所有格式属性
/// - `xf_type`: XF类型（单元格格式或样式格式）
pub struct XFRecord {
    xf: XF,
    xf_type: XFType,
}

impl XFRecord {
    pub fn new(xf: XF) -> Self {
        XFRecord {
            xf,
            xf_type: XFType::Cell,
        }
    }
}

impl Default for XFRecord {
    fn default() -> Self {
        XFRecord {
            xf: XF::default(),
            xf_type: XFType::Style,
        }
    }
}

impl BiffRecord for XFRecord {
    fn id(&self) -> u16 {
        0x00E0
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::new();

        // 2 bytes: font index
        buf.extend_from_slice(&self.xf.font_idx.to_le_bytes());

        // 2 bytes: format index
        buf.extend_from_slice(&self.xf.format_idx.to_le_bytes());

        // 2 bytes: protection + parent style
        let prot_bits: u16 = match self.xf_type {
            XFType::Cell => {
                ((self.xf.protection.cell_locked & 0x01) as u16)
                    | (((self.xf.protection.formula_hidden & 0x01) as u16) << 1)
            }
            XFType::Style => 0xFFF5,
        };
        let parent_idx: u16 = 0xFFFF;
        let prot_word = prot_bits | (parent_idx & 0xFFF0);
        buf.extend_from_slice(&prot_word.to_le_bytes());

        // 1 byte: alignment (horz + wrap + vert)
        let aln = ((self.xf.alignment.horz & 0x07) << 0)
            | ((self.xf.alignment.wrap & 0x01) << 3)
            | ((self.xf.alignment.vert & 0x07) << 4);
        buf.push(aln);

        // 1 byte: rotation
        buf.push(self.xf.alignment.rota);

        // 1 byte: indent + shrink + merge + direction
        let txt = ((self.xf.alignment.inde & 0x0F) << 0)
            | ((self.xf.alignment.shri & 0x01) << 4)
            | ((self.xf.alignment.merg & 0x01) << 5)
            | ((self.xf.alignment.dire & 0x03) << 6);
        buf.push(txt);

        // 1 byte: used attributes
        let used_attr = match self.xf_type {
            XFType::Cell => 0xF8,
            XFType::Style => 0xF4,
        };
        buf.push(used_attr);

        // Process borders (set colour to 0 if no line)
        let mut left_colour = self.xf.borders.left_colour;
        let mut right_colour = self.xf.borders.right_colour;
        let mut top_colour = self.xf.borders.top_colour;
        let mut bottom_colour = self.xf.borders.bottom_colour;
        let mut diag_colour = self.xf.borders.diag_colour;

        if self.xf.borders.left == 0 {
            left_colour = 0;
        }
        if self.xf.borders.right == 0 {
            right_colour = 0;
        }
        if self.xf.borders.top == 0 {
            top_colour = 0;
        }
        if self.xf.borders.bottom == 0 {
            bottom_colour = 0;
        }
        if self.xf.borders.diag == 0 {
            diag_colour = 0;
        }

        // 4 bytes: border lines (left, right, top, bottom + colors)
        let brd1: u32 = ((self.xf.borders.left as u32) & 0x0F)
            | (((self.xf.borders.right as u32) & 0x0F) << 4)
            | (((self.xf.borders.top as u32) & 0x0F) << 8)
            | (((self.xf.borders.bottom as u32) & 0x0F) << 12)
            | (((left_colour as u32) & 0x7F) << 16)
            | (((right_colour as u32) & 0x7F) << 23)
            | (((self.xf.borders.need_diag1 as u32) & 0x01) << 30)
            | (((self.xf.borders.need_diag2 as u32) & 0x01) << 31);
        buf.extend_from_slice(&brd1.to_le_bytes());

        // 4 bytes: border colors (top, bottom, diag + style)
        let brd2: u32 = ((top_colour as u32) & 0x7F)
            | (((bottom_colour as u32) & 0x7F) << 7)
            | (((diag_colour as u32) & 0x7F) << 14)
            | (((self.xf.borders.diag as u32) & 0x0F) << 21)
            | (((self.xf.pattern.pattern as u32) & 0x3F) << 26);
        buf.extend_from_slice(&brd2.to_le_bytes());

        // 2 bytes: pattern (fore + back color)
        let pat: u16 = ((self.xf.pattern.pattern_fore_colour as u16) & 0x7F)
            | (((self.xf.pattern.pattern_back_colour as u16) & 0x7F) << 7);
        buf.extend_from_slice(&pat.to_le_bytes());

        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_xf_record_id() {
        let record = XFRecord::new(XF::default());
        assert_eq!(record.id(), 0x00E0);
    }

    #[test]
    fn test_xf_record_data_size() {
        let record = XFRecord::new(XF::default());
        assert_eq!(record.data().len(), 20);
    }

    #[test]
    fn test_xf_record_style_type() {
        let xf = XF::default();
        let record = XFRecord {
            xf,
            xf_type: XFType::Style,
        };
        let data = record.data();

        // For style XF, protection word should be 0xFFF5
        assert_eq!(&data[4..6], &0xFFF5u16.to_le_bytes());
    }

    #[test]
    fn test_xf_record_cell_type() {
        let record = XFRecord::new(XF::default());
        let data = record.data();

        // For cell XF, used_attr should be 0xF8
        assert_eq!(data[9], 0xF8);
    }
}
