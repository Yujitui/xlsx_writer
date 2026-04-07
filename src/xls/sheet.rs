/// 代表 Excel 工作表中的一个单元格。
///
/// 此枚举封装了单元格可以包含的不同数据类型。
/// 每个变体对应于 Excel 单元格中存在的特定数据类型。
///
/// # 变体 (Variants)
///
/// * `Number(f64)` - 一个数值。Excel 内部将数字存储为 64 位浮点数。
///                   这包括整数、小数以及日期（以序列号形式存储）。
/// * `Text(String)` - 一串字符。这可以是纯文本，也可以是公式计算结果为文本的情况。
/// * `Boolean(bool)` - 一个逻辑值，`true` 或 `false`。
#[derive(Debug, Clone, PartialEq)]
pub enum XlsCell {
    /// 表示存储为 64 位浮点数的数值。
    /// 这是 Excel 中数字的标准内部表示形式。
    Number(f64),

    /// 表示存储为 Rust `String` 类型的文本数据。
    /// 这包括标签、名称以及任何公式得出的结果文本。
    Text(String),

    /// 表示一个布尔逻辑值（`true` 或 `false`）。
    Boolean(bool),
}

/// 代表一个 Excel 工作表。
///
/// 此结构体封装了单个工作表的所有信息，包括其名称和包含的单元格数据。
/// 数据以二维向量的形式存储，模拟了工作表的行列结构。
/// `rows` 向量的每个元素代表一行，行内的 `Vec<Option<XlsCell>>` 代表该行的各个单元格。
/// 使用 `Option` 是因为并非所有行列交叉点都一定有数据，`None` 表示该单元格为空。
#[derive(Debug)]
pub struct XlsSheet {
    /// 工作表的名称，例如 "Sheet1" 或用户自定义的名称。
    pub sheet_name: String,
    /// 存储工作表所有数据的二维向量。
    /// 外层向量的索引对应行号（从 0 开始）。
    /// 内层向量的索引对应列号（从 0 开始）。
    /// `Option<XlsCell>` 用于表示单元格是否有值：
    /// - `Some(cell)` 表示该单元格包含数据。
    /// - `None` 表示该单元格是空的。
    pub rows: Vec<Vec<Option<XlsCell>>>,
}

impl XlsSheet {
    
    pub fn new(sheet_name: String) -> XlsSheet {
        XlsSheet {
            sheet_name,
            rows: Vec::new(),
        }
    }
    
    /// 获取工作表中包含数据的实际范围。
    ///
    /// 此函数分析 `rows` 向量，找出包含至少一个非空单元格的最后行和最后列，
    /// 从而确定有效数据区域的边界。
    ///
    /// # 返回值
    ///
    /// * `Some((max_row, max_col))` - 如果工作表包含数据，则返回一个元组，
    ///   其中 `max_row` 是最后一行的索引（从 0 开始），
    ///   `max_col` 是最后一列的索引（从 0 开始）。
    /// * `None` - 如果工作表没有任何数据（即 `rows` 为空或所有单元格都为空）。
    pub fn data_range(&self) -> Option<(usize, usize)> {
        if self.rows.is_empty() {
            return None;
        }

        let max_row = self.rows.len() - 1;
        let max_col = self.rows.iter()
            .map(|row| row.len())
            .max()
            .unwrap_or(0);

        if max_col == 0 {
            None
        } else {
            Some((max_row, max_col - 1)) // 0-indexed
        }
    }

    /// 为遍历工作表中的所有非空单元格提供一个迭代器。
    ///
    /// 此函数返回一个迭代器，该迭代器会按行优先的顺序（从上到下，从左到右）
    /// 产出每一个包含数据的单元格及其对应的行列坐标。
    ///
    /// # 返回值
    ///
    /// 返回一个迭代器，其项目（Item）类型为 `(usize, usize, &XlsCell)`：
    /// - 第一个 `usize` 是行索引（从 0 开始）。
    /// - 第二个 `usize` 是列索引（从 0 开始）。
    /// - `&XlsCell` 是指向单元格数据的引用。
    pub fn cell_iterator(&self) -> impl Iterator<Item = (usize, usize, &XlsCell)> {
        self.rows.iter().enumerate().flat_map(|(row_idx, row)| {
            row.iter().enumerate().filter_map(move |(col_idx, cell)| {
                cell.as_ref().map(|c| (row_idx, col_idx, c))
            })
        })
    }

    /// 安全地在指定行列位置设置单元格值
    ///
    /// # 参数
    /// * `row` - 行索引（从 0 开始）
    /// * `col` - 列索引（从 0 开始）
    /// * `cell` - 要设置的单元格内容
    pub fn set_cell(&mut self, row: usize, col: usize, cell: XlsCell) {
        // 自动扩展行数（如果需要）
        if self.rows.len() <= row {
            self.rows.resize_with(row + 1, || vec![]);
        }

        // 自动扩展列数（如果需要）
        if self.rows[row].len() <= col {
            self.rows[row].resize_with(col + 1, || None);
        }

        // 设置单元格值
        self.rows[row][col] = Some(cell);
    }
}