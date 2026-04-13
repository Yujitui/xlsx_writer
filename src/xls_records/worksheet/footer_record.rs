use crate::xls_records::{utils::encode_biff_string_v2, BiffRecord};

/// FooterRecord 记录
///
/// 作用：存储页面打印的页脚内容
///
/// FooterRecord是Excel BIFF格式中的页脚记录（ID: 0x0015），用于定义打印页面时
/// 显示在页面底部的页脚文本。
///
/// ## 参数说明
///
/// - `s`: 页脚字符串，支持特殊标记：
///   - `&P`: 当前页码
///   - `&N`: 总页数
///   - `&D`: 当前日期
///   - `&T`: 当前时间
///   - `&F`: 文件名
///   - `&A`: 工作表名
/// - 默认值为 `&F`（显示文件名）
#[derive(Debug)]
pub struct FooterRecord {
    data: Vec<u8>,
}

impl FooterRecord {
    pub fn new(s: &str) -> Self {
        FooterRecord {
            data: encode_biff_string_v2(s),
        }
    }
}

impl Default for FooterRecord {
    fn default() -> Self {
        FooterRecord {
            data: encode_biff_string_v2("&F"),
        }
    }
}

impl BiffRecord for FooterRecord {
    fn id(&self) -> u16 {
        0x0015
    }

    fn data(&self) -> Vec<u8> {
        self.data.clone()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_footer_record_id() {
        let record = FooterRecord::default();
        assert_eq!(record.id(), 0x0015);
    }

    #[test]
    fn test_footer_record_default() {
        let record = FooterRecord::default();
        let data = record.data();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_footer_record_with_string() {
        let record = FooterRecord::new("Page &P");
        let data = record.data();
        assert!(!data.is_empty());
    }
}
