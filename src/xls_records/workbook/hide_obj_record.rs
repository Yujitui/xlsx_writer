use super::BiffRecord;

/// HideObjRecord 记录
///
/// 作用：控制隐藏对象的显示方式
///
/// HideObjRecord是Excel BIFF格式中的隐藏对象记录（ID: 0x008D），用于定义
/// 如何显示工作表中隐藏的对象（如图表、图形等）。
///
/// ## 参数说明
///
/// - 固定值：0x0000（默认，显示所有对象）
#[derive(Debug, Clone, Default)]
pub struct HideObjRecord;

impl HideObjRecord {
    pub fn new() -> Self {
        Self
    }
}

impl BiffRecord for HideObjRecord {
    fn id(&self) -> u16 {
        0x008D
    }

    fn data(&self) -> Vec<u8> {
        vec![0x00, 0x00]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hide_obj_record_id() {
        let record = HideObjRecord::new();
        assert_eq!(record.id(), 0x008D);
    }

    #[test]
    fn test_hide_obj_record_data_size() {
        let record = HideObjRecord::new();
        assert_eq!(record.data().len(), 2);
    }
}
