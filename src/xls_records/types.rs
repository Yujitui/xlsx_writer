use crate::cell::Cell;

/// 用于 XLS 解析的临时工作表结构
///
/// 此结构仅在解析 XLS 文件时使用，用于存储解析过程中的单元格数据。
/// 解析完成后将转换为 WorkSheet。
#[derive(Debug)]
pub struct XlsSheet {
    /// 工作表名称
    pub sheet_name: String,
    /// 单元格数据（行优先）
    pub rows: Vec<Vec<Option<Cell>>>,
}

impl XlsSheet {
    /// 创建新的 XlsSheet
    pub fn new(sheet_name: String) -> Self {
        XlsSheet {
            sheet_name,
            rows: Vec::new(),
        }
    }

    /// 在指定位置设置单元格值
    pub fn set_cell(&mut self, row: usize, col: usize, cell: Cell) {
        // 自动扩展行数
        if self.rows.len() <= row {
            self.rows.resize_with(row + 1, || vec![]);
        }
        // 自动扩展列数
        if self.rows[row].len() <= col {
            self.rows[row].resize_with(col + 1, || None);
        }
        self.rows[row][col] = Some(cell);
    }
}
