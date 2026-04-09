use super::BiffRecord;

/// PrecisionRecord 记录
///
/// 作用：控制数值精度显示
///
/// PrecisionRecord是Excel BIFF格式中的精度记录（ID: 0x000E），用于定义
/// 工作簿中数值显示的精度（按显示值计算还是按存储值计算）。
///
/// ## 参数说明
///
/// - `use_real_values`: 精度选项
///   - true = 按照显示的精度计算公式结果（默认）
///   - false = 按照存储的完整精度计算
#[derive(Debug)]
pub struct PrecisionRecord {
    use_real_values: bool,
}

impl PrecisionRecord {
    pub fn new(use_real_values: bool) -> Self {
        PrecisionRecord { use_real_values }
    }
}

impl Default for PrecisionRecord {
    fn default() -> Self {
        PrecisionRecord::new(true)
    }
}

impl BiffRecord for PrecisionRecord {
    fn id(&self) -> u16 {
        0x000E
    }

    fn data(&self) -> Vec<u8> {
        let value: u16 = if self.use_real_values { 1 } else { 0 };
        value.to_le_bytes().to_vec()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_precision_record_id() {
        let record = PrecisionRecord::new(false);
        assert_eq!(record.id(), 0x000E);
    }

    #[test]
    fn test_precision_record_data_size() {
        let record = PrecisionRecord::new(false);
        assert_eq!(record.data().len(), 2);
    }

    #[test]
    fn test_precision_use_real_values_false() {
        let record = PrecisionRecord::new(false);
        assert_eq!(record.data(), vec![0x00, 0x00]);
    }

    #[test]
    fn test_precision_use_real_values_true() {
        let record = PrecisionRecord::new(true);
        assert_eq!(record.data(), vec![0x01, 0x00]);
    }
}
