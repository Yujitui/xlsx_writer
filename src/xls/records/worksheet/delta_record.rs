use crate::xls::records::BiffRecord;

/// DeltaRecord 记录
///
/// 作用：设置工作表的迭代计算精度
///
/// DeltaRecord是Excel BIFF格式中的迭代增量记录（ID: 0x010），用于定义在
/// 迭代计算中使用的收敛阈值。当两次迭代计算结果的差值小于此值时，停止迭代。
///
/// ## 参数说明
///
/// - `delta`: 迭代收敛阈值
///   - 默认值为 0.001
///   - 值越小，计算结果越精确，但迭代次数可能增加
#[derive(Debug)]
pub struct DeltaRecord {
    delta: f64,
}

impl DeltaRecord {
    pub fn new(delta: f64) -> Self {
        DeltaRecord { delta }
    }
}

impl Default for DeltaRecord {
    fn default() -> Self {
        DeltaRecord { delta: 0.001 }
    }
}

impl BiffRecord for DeltaRecord {
    fn id(&self) -> u16 {
        0x010
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(8);
        buf.extend_from_slice(&self.delta.to_le_bytes());
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_delta_record_id() {
        let record = DeltaRecord::default();
        assert_eq!(record.id(), 0x010);
    }

    #[test]
    fn test_delta_record_default() {
        let record = DeltaRecord::default();
        assert_eq!(record.delta, 0.001);
    }

    #[test]
    fn test_delta_record_data_size() {
        let record = DeltaRecord::default();
        assert_eq!(record.data().len(), 8);
    }
}
