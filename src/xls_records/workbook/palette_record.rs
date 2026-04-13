use super::BiffRecord;

const DEFAULT_PALETTE: [u32; 56] = [
    0x00000000, 0xFFFFFF00, 0xFF000000, 0x00FF0000, 0x0000FF00, 0xFFFF0000, 0xFF00FF00, 0x00FFFF00,
    0x80000000, 0x00800000, 0x00008000, 0x80800000, 0x80008000, 0x00808000, 0xC0C0C000, 0x80808000,
    0x9999FF00, 0x99336600, 0xFFFFCC00, 0xCCFFFF00, 0x66006600, 0xFF808000, 0x0066CC00, 0xCCCCFF00,
    0x00008000, 0xFF00FF00, 0xFFFF0000, 0x00FFFF00, 0x80008000, 0x80000000, 0x00808000, 0x0000FF00,
    0x00CCFF00, 0xCCFFFF00, 0xCCFFCC00, 0xFFFF9900, 0x99CCFF00, 0xFF99CC00, 0xCC99FF00, 0xFFCC9900,
    0x3366FF00, 0x33CCCC00, 0x99CC0000, 0xFFCC0000, 0xFF990000, 0xFF660000, 0x66669900, 0x96969600,
    0x00336600, 0x33996600, 0x00330000, 0x33330000, 0x99330000, 0x99336600, 0x33339900, 0x33333300,
];

/// PaletteRecord 记录
///
/// 作用：定义工作簿的自定义调色板
///
/// PaletteRecord是Excel BIFF格式中的调色板记录（ID: 0x0092），用于定义
/// 工作簿中使用的自定义颜色。Excel默认提供56种颜色，此记录允许修改部分颜色。
///
/// ## 参数说明
///
/// - `modifications`: 颜色修改向量
///   - 每个元素为 (palette_index, red, green, blue)
///   - palette_index: 调色板索引 (0-55)
///   - red, green, blue: RGB值 (0-255)
#[derive(Debug)]
pub struct PaletteRecord {
    colors: Vec<u32>,
}

impl PaletteRecord {
    pub fn new(modifications: Vec<(u8, u8, u8, u8)>) -> Self {
        let mut colors = DEFAULT_PALETTE.to_vec();
        for (palette_index, red, green, blue) in modifications {
            let color = (red as u32) << 24 | (green as u32) << 16 | (blue as u32) << 8;
            colors[palette_index as usize] = color;
        }
        PaletteRecord { colors }
    }
}

impl Default for PaletteRecord {
    fn default() -> Self {
        PaletteRecord {
            colors: DEFAULT_PALETTE.to_vec(),
        }
    }
}

impl BiffRecord for PaletteRecord {
    fn id(&self) -> u16 {
        0x0092
    }

    fn data(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(2 + 56 * 4);
        buf.extend_from_slice(&56u16.to_le_bytes());
        for &color in &self.colors {
            buf.extend_from_slice(&color.to_be_bytes());
        }
        buf
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_palette_record_id() {
        let record = PaletteRecord::default();
        assert_eq!(record.id(), 0x0092);
    }

    #[test]
    fn test_palette_record_data_size() {
        let record = PaletteRecord::default();
        assert_eq!(record.data().len(), 226);
    }

    #[test]
    fn test_palette_record_with_modifications() {
        let record = PaletteRecord::new(vec![(0, 255, 0, 0)]);
        assert_eq!(record.colors[0], 0xFF000000);
    }
}
