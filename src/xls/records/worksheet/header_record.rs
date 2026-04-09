use crate::xls::records::{utils::encode_biff_string_v2, BiffRecord};

/// HeaderRecord 记录
///
/// 作用：存储页面打印的页眉内容
///
/// HeaderRecord是Excel BIFF格式中的页眉记录（ID: 0x0014），用于定义打印页面时
/// 显示在页面顶部的页眉文本。
///
/// ## 参数说明
///
/// - `s`: 页眉字符串，支持特殊标记：
///   - `&P`: 当前页码
///   - `&N`: 总页数
///   - `&D`: 当前日期
///   - `&T`: 当前时间
///   - `&F`: 文件名
///   - `&A`: 工作表名
/// - 默认值为 `&P`（显示页码）
#[derive(Debug)]
pub struct HeaderRecord {
    data: Vec<u8>,
}

impl HeaderRecord {
    pub fn new(s: &str) -> Self {
        HeaderRecord {
            data: encode_biff_string_v2(s),
        }
    }
}

impl Default for HeaderRecord {
    fn default() -> Self {
        HeaderRecord {
            data: encode_biff_string_v2("&P"),
        }
    }
}

impl BiffRecord for HeaderRecord {
    fn id(&self) -> u16 {
        0x0014
    }

    fn data(&self) -> Vec<u8> {
        self.data.clone()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_header_record_id() {
        let record = HeaderRecord::default();
        assert_eq!(record.id(), 0x0014);
    }

    #[test]
    fn test_header_record_default() {
        let record = HeaderRecord::default();
        let data = record.data();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_header_record_with_string() {
        let record = HeaderRecord::new("Page &P of &N");
        let data = record.data();
        assert!(!data.is_empty());
    }
}
