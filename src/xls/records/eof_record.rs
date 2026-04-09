use super::BiffRecord;

/// EOF (End of File) 记录
///
/// ## 作用
///
/// 标识BIFF段落的结束。每个BIFF段（如工作簿globals或worksheet）结束时需要此记录。
/// EOF记录没有数据部分，data()返回空向量。
#[derive(Debug, Default)]
pub struct EofRecord;

impl EofRecord {
    pub fn new() -> Self {
        EofRecord
    }
}

impl BiffRecord for EofRecord {
    fn id(&self) -> u16 {
        0x000A
    }

    fn data(&self) -> Vec<u8> {
        Vec::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_eof_record_id() {
        let record = EofRecord::new();
        assert_eq!(record.id(), 0x000A);
    }

    #[test]
    fn test_eof_record_data_size() {
        let record = EofRecord::new();
        assert_eq!(record.data().len(), 0);
    }
}
