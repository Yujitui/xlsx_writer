/// BIFF (Binary Interchange File Format) 记录类型枚举
/// 用于标识 Excel 97-2003 (.xls) 文件中的各种记录
#[allow(non_camel_case_types)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum RecordType {
    // 基础结构记录
    BOF = 0x0809,        // Beginning of File
    EOF = 0x000A,        // End of File
    BOUNDSHEET = 0x0085, // Sheet Information
    SST = 0x00FC,        // Shared String Table
    DIMENSIONS = 0x0200, // Sheet Dimensions
    Row = 0x0208,        // Row Information
    NUMBER = 0x0203,     // Floating Point Number
    RK = 0x027E,         // Encoded Integer/Float
    RKOLD = 0x007E,      // 旧版 RK 记录
    MULRK = 0x00BD,      // Multiple RK Records
    LABEL = 0x0204,      // In-place String Label
    LABELSST = 0x00FD,   // String Label from SST
    FORMULA = 0x0006,    // Formula
    BOOL = 0x0202,       // 假设的布尔值记录类型

    // 样式和格式记录
    FONT = 0x0031,    // Font description
    XF = 0x00E0,      // Extended Format
    STYLE = 0x0293,   // Style information
    PALETTE = 0x0092, // Palette for custom colors
    FORMAT = 0x041E,

    // 窗口和显示记录
    WINDOW1 = 0x003D,     // Window settings
    WINDOW2 = 0x023E,     // Sheet window settings
    BACKUP = 0x0040,      // Save backup flag
    HIDEOBJ = 0x008D,     // Object visibility
    SETUP = 0x00A1,       // Page setup
    DEFCOLWIDTH = 0x0055, // Default column width
    COLINFO = 0x007D,     // Column information

    // 打印和布局记录
    PRINTHEADERS = 0x002A,   // Print row/column headers
    PRINTGRIDLINES = 0x002B, // Print gridlines
    GRIDSET = 0x0082,        // Grid settings
    GUTS = 0x0080,           // Size of row/column headings

    // 保护和安全记录
    PROTECT = 0x0012,     // Protection flag
    SCENPROTECT = 0x00DD, // Scenario protection
    OBJPROTECT = 0x0063,  // Object protection

    // 公式和名称记录
    NAME = 0x0018,        // Defined names
    EXTERNNAME = 0x0023,  // External defined names
    EXTERNSHEET = 0x0017, // External sheet references

    // 对象和图表记录
    MSO_DRAWING = 0x00EC, // Microsoft Office drawing
    OBJ = 0x005D,         // Object descriptor
    TXO = 0x01B6,         // Text object
    NOTE = 0x001C,        // Cell comment

    // 数据记录
    ARRAY = 0x0221,   // Array formula
    SHRFMLA = 0x04BC, // Shared formula
    TABLE = 0x0236,   // Data table

    // 其他功能记录
    WRITEACCESS = 0x005C,    // Write access user name
    WSBOOL = 0x0081,         // Worksheet options
    MMS = 0x00C1,            // Multiplan menu structure
    ADDMENU = 0x00C2,        // Menu addition
    DELMENU = 0x00BF,        // Menu deletion
    TOOLBARHDR = 0x0161,     // Toolbar header
    TOOLBAREND = 0x0162,     // Toolbar end
    BUTTONPROPERTY = 0x013D, // Button property
    MERGECELLS = 0x00E5,     // Merged cells
    UNCALCED = 0x005E,       // Uncalculated formula flag

    Unknown = 0x0000,
}

impl RecordType {
    /// Converts a raw u16 opcode to a RecordType variant.
    pub fn from_u16(code: u16) -> Self {
        match code {
            0x0809 => RecordType::BOF,
            0x000A => RecordType::EOF,
            0x0085 => RecordType::BOUNDSHEET,
            0x00FC => RecordType::SST,
            0x0200 => RecordType::DIMENSIONS,
            0x0208 => RecordType::Row,
            0x0203 => RecordType::NUMBER,
            0x027E => RecordType::RK,
            0x007E => RecordType::RKOLD,
            0x00BD => RecordType::MULRK,
            0x0204 => RecordType::LABEL,
            0x00FD => RecordType::LABELSST,
            0x0006 => RecordType::FORMULA,
            0x0031 => RecordType::FONT,
            0x00E0 => RecordType::XF,
            0x041E => RecordType::FORMAT,
            0x003D => RecordType::WINDOW1,
            0x0293 => RecordType::STYLE,
            0x023E => RecordType::WINDOW2,
            0x0040 => RecordType::BACKUP,
            0x008D => RecordType::HIDEOBJ,
            0x00A1 => RecordType::SETUP,
            0x0055 => RecordType::DEFCOLWIDTH,
            0x007D => RecordType::COLINFO,
            0x002A => RecordType::PRINTHEADERS,
            0x002B => RecordType::PRINTGRIDLINES,
            0x0082 => RecordType::GRIDSET,
            0x0080 => RecordType::GUTS,
            0x0012 => RecordType::PROTECT,
            0x00DD => RecordType::SCENPROTECT,
            0x0063 => RecordType::OBJPROTECT,
            0x0018 => RecordType::NAME,
            0x0023 => RecordType::EXTERNNAME,
            0x0017 => RecordType::EXTERNSHEET,
            0x00EC => RecordType::MSO_DRAWING,
            0x005D => RecordType::OBJ,
            0x01B6 => RecordType::TXO,
            0x001C => RecordType::NOTE,
            0x0221 => RecordType::ARRAY,
            0x04BC => RecordType::SHRFMLA,
            0x0236 => RecordType::TABLE,
            0x005C => RecordType::WRITEACCESS,
            0x0081 => RecordType::WSBOOL,
            0x00C1 => RecordType::MMS,
            0x00C2 => RecordType::ADDMENU,
            0x00BF => RecordType::DELMENU,
            0x0161 => RecordType::TOOLBARHDR,
            0x0162 => RecordType::TOOLBAREND,
            0x013D => RecordType::BUTTONPROPERTY,
            0x00E5 => RecordType::MERGECELLS,
            0x005E => RecordType::UNCALCED,
            _ => RecordType::Unknown,
        }
    }

    /// 获取记录类型的 u16 值
    pub fn to_u16(self) -> u16 {
        self as u16
    }

    pub fn to_string(&self) -> String {
        match self {
            RecordType::BOF => "BOF (Beginning of File)".to_string(),
            RecordType::EOF => "EOF (End of File)".to_string(),
            RecordType::BOUNDSHEET => "BoundSheet (Sheet Information)".to_string(),
            RecordType::SST => "SST (Shared String Table)".to_string(),
            RecordType::DIMENSIONS => "Dimensions (Sheet Dimensions)".to_string(),
            RecordType::Row => "Row (Row Information)".to_string(),
            RecordType::NUMBER => "Number (Floating Point Number)".to_string(),
            RecordType::RK => "RK (Encoded Integer/Float)".to_string(),
            RecordType::RKOLD => "RKOld (Older version of RK)".to_string(),
            RecordType::MULRK => "MulRK (Multiple RK Records)".to_string(),
            RecordType::LABEL => "Label (In-place String Label)".to_string(),
            RecordType::LABELSST => "LabelSST (String Label from SST)".to_string(),
            RecordType::FORMULA => "Formula".to_string(),
            RecordType::FONT => "FONT (Font description)".to_string(),
            RecordType::FORMAT => "Format".to_string(),
            RecordType::XF => "XF (Extended Format)".to_string(),
            RecordType::STYLE => "STYLE (Style information)".to_string(),
            RecordType::WINDOW1 => "WINDOW1 (Window settings)".to_string(),
            RecordType::WINDOW2 => "WINDOW2 (Sheet window settings)".to_string(),
            RecordType::BACKUP => "BACKUP (Save backup flag)".to_string(),
            RecordType::HIDEOBJ => "HIDEOBJ (Object visibility)".to_string(),
            RecordType::SETUP => "SETUP (Page setup)".to_string(),
            RecordType::DEFCOLWIDTH => "DEFCOLWIDTH (Default column width)".to_string(),
            RecordType::COLINFO => "COLINFO (Column information)".to_string(),
            RecordType::PRINTHEADERS => "PRINTHEADERS (Print row/column headers)".to_string(),
            RecordType::PRINTGRIDLINES => "PRINTGRIDLINES (Print gridlines)".to_string(),
            RecordType::GRIDSET => "GRIDSET (Grid settings)".to_string(),
            RecordType::GUTS => "GUTS (Size of row/column headings)".to_string(),
            RecordType::PROTECT => "PROTECT (Protection flag)".to_string(),
            RecordType::SCENPROTECT => "SCENPROTECT (Scenario protection)".to_string(),
            RecordType::OBJPROTECT => "OBJPROTECT (Object protection)".to_string(),
            RecordType::NAME => "NAME (Defined names)".to_string(),
            RecordType::EXTERNNAME => "EXTERNNAME (External defined names)".to_string(),
            RecordType::EXTERNSHEET => "EXTERNSHEET (External sheet references)".to_string(),
            RecordType::MSO_DRAWING => "MSO_DRAWING (Microsoft Office drawing)".to_string(),
            RecordType::OBJ => "OBJ (Object descriptor)".to_string(),
            RecordType::TXO => "TXO (Text object)".to_string(),
            RecordType::NOTE => "NOTE (Cell comment)".to_string(),
            RecordType::ARRAY => "ARRAY (Array formula)".to_string(),
            RecordType::SHRFMLA => "SHRFMLA (Shared formula)".to_string(),
            RecordType::TABLE => "TABLE (Data table)".to_string(),
            RecordType::WRITEACCESS => "WRITEACCESS (Write access user name)".to_string(),
            RecordType::WSBOOL => "WSBOOL (Worksheet options)".to_string(),
            RecordType::MMS => "MMS (Multiplan menu structure)".to_string(),
            RecordType::ADDMENU => "ADDMENU (Menu addition)".to_string(),
            RecordType::DELMENU => "DELMENU (Menu deletion)".to_string(),
            RecordType::TOOLBARHDR => "TOOLBARHDR (Toolbar header)".to_string(),
            RecordType::TOOLBAREND => "TOOLBAREND (Toolbar end)".to_string(),
            RecordType::BUTTONPROPERTY => "BUTTONPROPERTY (Button property)".to_string(),
            RecordType::MERGECELLS => "MERGECELLS (Merged cells)".to_string(),
            RecordType::UNCALCED => "UNCALCED (Uncalculated formula flag)".to_string(),
            RecordType::Unknown => "Unknown".to_string(),
            _ => "Unknown".to_string(),
        }
    }
}
