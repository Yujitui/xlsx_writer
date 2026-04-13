use serde::Deserialize;
use std::hash::Hash;

/// 单个工作表的读取配置
///
/// 该结构体用于定义读取Excel工作表时的各种参数配置，
/// 包括工作表名称、强制字符串列和需要跳过的行数等。
///
/// # 字段说明
///
/// - `sheet_name`: 指定要读取的工作表名称
/// - `force_string_cols`: 可选的列名列表，指定哪些列应强制作为字符串类型读取
/// - `skip_rows`: 可选的行数，指定在读取数据前需要跳过的行数
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "snake_case")]
pub struct ReadSheet {
    /// 工作表名称
    ///
    /// 必需字段，指定要读取的具体工作表名称
    pub sheet_name: String,

    /// 强制将指定列作为字符串读取
    ///
    /// 可选字段，在JSON中表示为字符串数组。
    /// 如果省略此字段，则为null（Rust中的None）
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub force_string_cols: Option<Vec<String>>,

    /// 跳过的行数
    ///
    /// 可选字段，在JSON中表示为整数。
    /// 如果省略此字段，则为null（Rust中的None）
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub skip_rows: Option<usize>,
}

impl ReadSheet {
    /// 创建新的工作表读取配置
    ///
    /// 构造一个基本的`ReadSheet`实例，使用指定的工作表名称，
    /// 其他配置项（强制字符串列和跳行数）初始化为`None`。
    ///
    /// # 参数
    ///
    /// * `sheet_name` - 要读取的工作表名称字符串
    ///
    /// # 返回值
    ///
    /// 返回一个新的`ReadSheet`实例，所有可选配置项初始为`None`
    pub fn new(sheet_name: String) -> ReadSheet {
        ReadSheet {
            sheet_name,
            ..Default::default()
        }
    }
}

// 为 ReadSheet 实现 Default trait，便于测试和默认配置
impl Default for ReadSheet {
    fn default() -> Self {
        ReadSheet {
            sheet_name: "Sheet1".to_string(),
            force_string_cols: None,
            skip_rows: None,
        }
    }
}

impl Hash for ReadSheet {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.sheet_name.hash(state);
    }
}

impl PartialEq for ReadSheet {
    fn eq(&self, other: &Self) -> bool {
        self.sheet_name == other.sheet_name
    }
}

impl Eq for ReadSheet {}
