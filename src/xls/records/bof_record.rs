use crate::xls::records::BiffRecord;

/// BOF记录类型
///
/// ## 作用
///
/// 定义BOF记录可以标识的BIFF段落类型
///
/// ## 参数说明
///
/// - `WorkbookGlobals` (0x0005): 工作簿全局信息（包含所有工作表的共享数据，如SST、XF、Font等）
/// - `VisualBasicModule` (0x0006): Visual Basic模块
/// - `Worksheet` (0x0010): 工作表
/// - `Chart` (0x0020): 图表
/// - `MacroSheet` (0x0040): 宏表
/// - `Workspace` (0x0100): 工作区
#[repr(u16)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BofType {
    /// 工作簿全局信息（包含所有工作表的共享数据，如SST、XF、Font等）
    WorkbookGlobals = 0x0005,
    /// Visual Basic模块
    VisualBasicModule = 0x0006,
    /// 工作表
    Worksheet = 0x0010,
    /// 图表
    Chart = 0x0020,
    /// 宏表
    MacroSheet = 0x0040,
    /// 工作区
    Workspace = 0x0100,
}

impl BofType {
    pub fn to_u16(&self) -> u16 {
        *self as u16
    }
}

/// BOF (Begin of File) 记录
///
/// ## 作用
///
/// 标识一个新的BIFF段落的开始。BIFF文件由多个段落组成，每个段落以BOF开始，以EOF结束。
/// 例如：工作簿globals段落、工作表段落、图表段落等。
///
/// ## 参数说明
///
/// - `bof_type`: BOF记录类型，指定此段落的内容类型（见BofType枚举）
///   - 可选值：BofType::WorkbookGlobals, BofType::Worksheet, BofType::Chart 等
#[derive(Debug)]
pub struct BoFRecord {
    /// BOF记录类型，指定此段落的内容类型
    pub bof_type: BofType,
}

impl BoFRecord {
    pub fn new(bof_type: BofType) -> Self {
        BoFRecord { bof_type }
    }
}

impl BiffRecord for BoFRecord {
    fn id(&self) -> u16 {
        0x0809 // BOF record ID
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(10);
        buf.extend_from_slice(&0x0600u16.to_le_bytes()); // Version
        buf.extend_from_slice(&(self.bof_type as u16).to_le_bytes()); // Type
        buf.extend_from_slice(&0x0DBBu16.to_le_bytes()); // Build
        buf.extend_from_slice(&0x07CCu16.to_le_bytes()); // Year
        buf.extend_from_slice(&0x00u8.to_le_bytes()); // Flags
        buf.extend_from_slice(&0x06u8.to_le_bytes()); // VerCanRead
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_bof_record_id() {
        let record = BoFRecord::new(BofType::Worksheet);
        assert_eq!(record.id(), 0x0809);
    }

    #[test]
    fn test_bof_record_data_size() {
        let record = BoFRecord::new(BofType::Worksheet);
        assert_eq!(record.data().len(), 10);
    }

    #[test]
    fn test_bof_record_worksheet_type() {
        let record = BoFRecord::new(BofType::Worksheet);
        let data = record.data();
        assert_eq!(&data[2..4], &0x0010u16.to_le_bytes());
    }

    #[test]
    fn test_bof_record_workbook_type() {
        let record = BoFRecord::new(BofType::WorkbookGlobals);
        let data = record.data();
        assert_eq!(&data[2..4], &0x0005u16.to_le_bytes());
    }

    #[test]
    fn test_bof_type_to_u16() {
        assert_eq!(BofType::Worksheet.to_u16(), 0x0010);
        assert_eq!(BofType::WorkbookGlobals.to_u16(), 0x0005);
        assert_eq!(BofType::Chart.to_u16(), 0x0020);
    }
}
