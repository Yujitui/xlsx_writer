/// 将字符串编码为 BIFF 格式（1字节长度版本 - upack1）
/// 用于 BoundSheetRecord、FontRecord 等
pub fn encode_biff_string_v1(s: &str) -> Vec<u8> {
    if s.is_ascii() {
        let bytes = s.as_bytes();
        let mut result = Vec::with_capacity(2 + bytes.len());
        result.push(bytes.len() as u8);
        result.push(0x00); // flag = compressed ASCII
        result.extend_from_slice(bytes);
        result
    } else {
        let utf16: Vec<u8> = s.encode_utf16().flat_map(|c| c.to_le_bytes()).collect();
        let char_count = (utf16.len() / 2) as u8;
        let mut result = Vec::with_capacity(2 + utf16.len());
        result.push(char_count);
        result.push(0x01); // flag = uncompressed UTF-16
        result.extend_from_slice(&utf16);
        result
    }
}

/// 将字符串编码为 BIFF 格式（2字节长度版本 - upack2）
/// 用于 NumberFormatRecord 等
pub fn encode_biff_string_v2(s: &str) -> Vec<u8> {
    if s.is_ascii() {
        let bytes = s.as_bytes();
        let char_count = bytes.len() as u16;
        let mut result = Vec::with_capacity(3 + bytes.len());
        result.extend_from_slice(&char_count.to_le_bytes());
        result.push(0x00); // flag = compressed ASCII
        result.extend_from_slice(bytes);
        result
    } else {
        let utf16: Vec<u8> = s.encode_utf16().flat_map(|c| c.to_le_bytes()).collect();
        let char_count = (utf16.len() / 2) as u16;
        let mut result = Vec::with_capacity(3 + utf16.len());
        result.extend_from_slice(&char_count.to_le_bytes());
        result.push(0x01); // flag = uncompressed UTF-16
        result.extend_from_slice(&utf16);
        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_encode_biff_string_v1_ascii() {
        let result = encode_biff_string_v1("Sheet1");
        assert_eq!(result[0], 6); // length
        assert_eq!(result[1], 0x00); // flag = ASCII
        assert_eq!(&result[2..], b"Sheet1");
    }

    #[test]
    fn test_encode_biff_string_v1_unicode() {
        let result = encode_biff_string_v1("表格");
        assert_eq!(result[0], 2); // char count
        assert_eq!(result[1], 0x01); // flag = UTF-16
    }

    #[test]
    fn test_encode_biff_string_v2_ascii() {
        let result = encode_biff_string_v2("General");
        assert_eq!(&result[0..2], &7u16.to_le_bytes()); // length (2 bytes)
        assert_eq!(result[2], 0x00); // flag = ASCII
        assert_eq!(&result[3..], b"General");
    }

    #[test]
    fn test_encode_biff_string_v2_unicode() {
        let result = encode_biff_string_v2("格式");
        assert_eq!(&result[0..2], &2u16.to_le_bytes()); // char count (2 bytes)
        assert_eq!(result[2], 0x01); // flag = UTF-16
    }
}
