//! OLE Property Set Stream 构造器
//!
//! 用于生成 Excel .xls 文件中包含的两个元数据流：
//! - \x05SummaryInformation
//! - \x05DocumentSummaryInformation
//!
//! 格式参考 [MS-OLEPS] Object Linking and Embedding (OLE) Property Set Data Structures。

use std::time::{SystemTime, UNIX_EPOCH};

// ============================================================================
// 常量定义
// ============================================================================

/// 流固定大小：Excel 2004 (Mac) 对这两个流总是分配 8 个 512 字节扇区。
pub const OLE_METADATA_STREAM_SIZE: usize = 4096;

// VT 类型标识
const VT_I2: u32 = 0x0002;
const VT_I4: u32 = 0x0003;
const VT_BOOL: u32 = 0x000B;
const VT_LPSTR: u32 = 0x001E;
const VT_FILETIME: u32 = 0x0040;
const VT_VARIANT: u32 = 0x000C;
const VT_VECTOR: u32 = 0x1000;

/// SummaryInformation 的 FMTID (小端字节序)
const FMTID_SUMMARY_INFORMATION: [u8; 16] = [
    0xe0, 0x85, 0x9f, 0xf2, 0xf9, 0x4f, 0x68, 0x10,
    0xab, 0x91, 0x08, 0x00, 0x2b, 0x27, 0xb3, 0xd9,
];

/// DocumentSummaryInformation 第一节的 FMTID (小端字节序)
const FMTID_DOC_SUMMARY_INFORMATION: [u8; 16] = [
    0x02, 0xd5, 0xcd, 0xd5, 0x9c, 0x2e, 0x1b, 0x10,
    0x93, 0x97, 0x08, 0x00, 0x2b, 0x2c, 0xf9, 0xae,
];

// ============================================================================
// 公共数据结构
// ============================================================================

/// SummaryInformation 流属性
#[derive(Debug, Clone)]
pub struct SummaryInfoProps {
    /// codepage，Excel Mac 中文版使用 10008
    pub codepage: i16,
    /// 最后保存者（已按目标 codepage 编码的字节，或空）
    pub last_author: Vec<u8>,
    /// 创建应用程序名
    pub application_name: String,
    /// 文档安全级别
    pub security: i32,
    /// 最后保存时间
    pub last_saved_time: SystemTime,
}

impl Default for SummaryInfoProps {
    fn default() -> Self {
        Self {
            codepage: 10008,
            last_author: Vec::new(),
            application_name: "Microsoft Macintosh Excel".to_string(),
            security: 0,
            last_saved_time: SystemTime::now(),
        }
    }
}

/// DocumentSummaryInformation 流属性
///
/// 注意：heading_pairs 中的名称和 titles_of_parts 中的标题需要已经按目标
/// codepage 编码为字节（Excel Mac 中文版实际使用 GBK 编码）。
#[derive(Debug, Clone)]
pub struct DocSummaryInfoProps {
    /// codepage
    pub codepage: i16,
    /// 是否缩放裁剪缩略图
    pub scale_crop: bool,
    /// HeadingPairs：工作表类型名与数量的配对，如 [(b"工作表", 1)]
    pub heading_pairs: Vec<(Vec<u8>, i32)>,
    /// 每个部分的标题，通常为工作表名列表
    pub titles_of_parts: Vec<Vec<u8>>,
    /// 链接是否脏
    pub links_dirty: bool,
    /// 是否共享文档
    pub shared_doc: bool,
    /// 超链接是否已更改
    pub hlinks_changed: bool,
    /// 应用程序版本号，Excel Mac 写入 1048576
    pub version: i32,
}

impl Default for DocSummaryInfoProps {
    fn default() -> Self {
        Self {
            codepage: 10008,
            scale_crop: false,
            heading_pairs: vec![("工作表".as_bytes().to_vec(), 1)],
            titles_of_parts: vec!["Sheet1".as_bytes().to_vec()],
            links_dirty: false,
            shared_doc: false,
            hlinks_changed: false,
            version: 1048576,
        }
    }
}

// ============================================================================
// 内部构造器
// ============================================================================

/// 单个属性的序列化表示
#[derive(Debug)]
struct Property {
    id: u32,
    value: PropertyValue,
}

#[derive(Debug)]
enum PropertyValue {
    I2(i16),
    I4(i32),
    Bool(bool),
    Lpstr(Vec<u8>),
    Filetime(u64),
    VectorLpstr(Vec<Vec<u8>>),
    VectorVariant(Vec<Variant>),
}

#[derive(Debug)]
enum Variant {
    Lpstr(Vec<u8>),
    I4(i32),
}

/// OLEPS Section 构造器
struct SectionBuilder {
    fmtid: [u8; 16],
    properties: Vec<Property>,
}

impl SectionBuilder {
    fn new(fmtid: [u8; 16]) -> Self {
        Self {
            fmtid,
            properties: Vec::new(),
        }
    }

    fn add_i2(mut self, id: u32, value: i16) -> Self {
        self.properties.push(Property { id, value: PropertyValue::I2(value) });
        self
    }

    fn add_i4(mut self, id: u32, value: i32) -> Self {
        self.properties.push(Property { id, value: PropertyValue::I4(value) });
        self
    }

    fn add_bool(mut self, id: u32, value: bool) -> Self {
        self.properties.push(Property { id, value: PropertyValue::Bool(value) });
        self
    }

    fn add_lpstr(mut self, id: u32, bytes: Vec<u8>) -> Self {
        self.properties.push(Property { id, value: PropertyValue::Lpstr(bytes) });
        self
    }

    fn add_filetime(mut self, id: u32, value: u64) -> Self {
        self.properties.push(Property { id, value: PropertyValue::Filetime(value) });
        self
    }

    fn add_vector_lpstr(mut self, id: u32, values: Vec<Vec<u8>>) -> Self {
        self.properties.push(Property { id, value: PropertyValue::VectorLpstr(values) });
        self
    }

    fn add_vector_variant(mut self, id: u32, values: Vec<Variant>) -> Self {
        self.properties.push(Property { id, value: PropertyValue::VectorVariant(values) });
        self
    }

    /// 构建整个 Property Set Stream，并补齐到 target_size。
    fn build(self, target_size: usize) -> Vec<u8> {
        // 1. 先序列化每个属性值，得到其字节和大小
        let mut value_bytes: Vec<Vec<u8>> = Vec::with_capacity(self.properties.len());
        for prop in &self.properties {
            value_bytes.push(serialize_property_value(&prop.value));
        }

        // 2. 计算属性表和属性值的布局
        // 属性表起始偏移 = 8（SectionSize + NumProperties）
        let num_props = self.properties.len();
        let table_size = 8 * num_props;
        let values_base: u32 = 8 + table_size as u32;

        let mut table = Vec::with_capacity(table_size);
        let mut values = Vec::new();
        let mut current_offset = values_base;

        for (prop, bytes) in self.properties.iter().zip(value_bytes.iter()) {
            // 属性表条目：PID (4) + offset (4)
            table.extend_from_slice(&prop.id.to_le_bytes());
            table.extend_from_slice(&current_offset.to_le_bytes());
            values.extend_from_slice(bytes);
            current_offset += bytes.len() as u32;
        }

        // 3. Section Data：size + num_props + table + values
        let mut section = Vec::new();
        let section_body_size = 4 + 4 + table.len() + values.len();
        // Excel 的 section size 按 16 字节对齐
        let section_size = align_up(section_body_size, 16);
        section.extend_from_slice(&(section_size as u32).to_le_bytes());
        section.extend_from_slice(&(num_props as u32).to_le_bytes());
        section.extend_from_slice(&table);
        section.extend_from_slice(&values);
        section.resize(section_size, 0);

        // 4. Property Set Stream Header + Section Header
        let section_offset: u32 = 48; // 28 字节 header + 20 字节 section header
        let mut stream = Vec::with_capacity(target_size);
        stream.extend_from_slice(&build_stream_header());
        stream.extend_from_slice(&self.fmtid);
        stream.extend_from_slice(&section_offset.to_le_bytes());
        stream.extend_from_slice(&section);

        // 5. 补齐到目标大小
        stream.resize(target_size, 0);
        stream
    }
}

// ============================================================================
// 序列化辅助函数
// ============================================================================

/// OLEPS Stream Header，固定 28 字节（与 Excel Mac 完全一致）
fn build_stream_header() -> [u8; 28] {
    [
        0xfe, 0xff, // ByteOrder = 0xFFFE (Little-Endian)
        0x00, 0x00, // 
        0x1a, 0x05, // 版本字段，与 Excel Mac 一致
        0x01, 0x00, //
        0x00, 0x00, 0x00, 0x00, // OSVersion = 0
        0x00, 0x00, 0x00, 0x00, // OS = 0
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // CLSID = empty
        0x01, 0x00, 0x00, 0x00, // NumSections = 1
    ]
}

fn serialize_property_value(value: &PropertyValue) -> Vec<u8> {
    let mut buf = Vec::new();
    match value {
        PropertyValue::I2(v) => {
            buf.extend_from_slice(&VT_I2.to_le_bytes());
            buf.extend_from_slice(&v.to_le_bytes());
            pad_to(&mut buf, 8);
        }
        PropertyValue::I4(v) => {
            buf.extend_from_slice(&VT_I4.to_le_bytes());
            buf.extend_from_slice(&v.to_le_bytes());
        }
        PropertyValue::Bool(v) => {
            buf.extend_from_slice(&VT_BOOL.to_le_bytes());
            let flag: u16 = if *v { 0xFFFF } else { 0x0000 };
            buf.extend_from_slice(&flag.to_le_bytes());
            pad_to(&mut buf, 8);
        }
        PropertyValue::Lpstr(bytes) => {
            buf.extend_from_slice(&VT_LPSTR.to_le_bytes());
            // 长度包含 null 终止符
            let len = if bytes.is_empty() { 1 } else { bytes.len() + if bytes.ends_with(&[0]) { 0 } else { 1 } };
            buf.extend_from_slice(&(len as u32).to_le_bytes());
            if bytes.is_empty() {
                buf.push(0);
            } else {
                buf.extend_from_slice(bytes);
                if !bytes.ends_with(&[0]) {
                    buf.push(0);
                }
            }
            pad_to(&mut buf, 4);
        }
        PropertyValue::Filetime(v) => {
            buf.extend_from_slice(&VT_FILETIME.to_le_bytes());
            buf.extend_from_slice(&v.to_le_bytes());
        }
        PropertyValue::VectorLpstr(items) => {
            buf.extend_from_slice(&(VT_VECTOR | VT_LPSTR).to_le_bytes());
            buf.extend_from_slice(&(items.len() as u32).to_le_bytes());
            for item in items {
                // Excel 在 vector 中的 LPSTR 项不补 4 字节对齐
                let len = if item.is_empty() { 1 } else { item.len() + if item.ends_with(&[0]) { 0 } else { 1 } };
                buf.extend_from_slice(&(len as u32).to_le_bytes());
                if item.is_empty() {
                    buf.push(0);
                } else {
                    buf.extend_from_slice(item);
                    if !item.ends_with(&[0]) {
                        buf.push(0);
                    }
                }
            }
        }
        PropertyValue::VectorVariant(variants) => {
            buf.extend_from_slice(&(VT_VECTOR | VT_VARIANT).to_le_bytes());
            buf.extend_from_slice(&(variants.len() as u32).to_le_bytes());
            for variant in variants {
                match variant {
                    Variant::Lpstr(bytes) => {
                        buf.extend_from_slice(&VT_LPSTR.to_le_bytes());
                        let len = if bytes.is_empty() { 1 } else { bytes.len() + if bytes.ends_with(&[0]) { 0 } else { 1 } };
                        buf.extend_from_slice(&(len as u32).to_le_bytes());
                        if bytes.is_empty() {
                            buf.push(0);
                        } else {
                            buf.extend_from_slice(bytes);
                            if !bytes.ends_with(&[0]) {
                                buf.push(0);
                            }
                        }
                    }
                    Variant::I4(v) => {
                        buf.extend_from_slice(&VT_I4.to_le_bytes());
                        buf.extend_from_slice(&v.to_le_bytes());
                    }
                }
            }
            // Excel 中该属性末尾会补到 4 字节边界
            pad_to(&mut buf, 4);
        }
    }
    buf
}

fn pad_to(buf: &mut Vec<u8>, alignment: usize) {
    let rem = buf.len() % alignment;
    if rem != 0 {
        buf.resize(buf.len() + alignment - rem, 0);
    }
}

fn align_up(size: usize, alignment: usize) -> usize {
    let rem = size % alignment;
    if rem == 0 { size } else { size + alignment - rem }
}

/// 将 SystemTime 转换为 FILETIME（自 1601-01-01 以来的 100ns 间隔数）
fn system_time_to_filetime(st: SystemTime) -> u64 {
    let duration = st.duration_since(UNIX_EPOCH).unwrap_or_default();
    let seconds = duration.as_secs();
    let nanos = duration.subsec_nanos() as u64;
    // 11644473600 = 1601-01-01 到 1970-01-01 的秒数
    (seconds + 11644473600) * 10_000_000 + nanos / 100
}

// ============================================================================
// 公共 API
// ============================================================================

/// 构造 SummaryInformation 流（固定 4096 字节）
pub fn build_summary_information_stream(props: &SummaryInfoProps) -> Vec<u8> {
    SectionBuilder::new(FMTID_SUMMARY_INFORMATION)
        .add_i2(0x01, props.codepage)
        .add_lpstr(0x08, props.last_author.clone())
        .add_filetime(0x0D, system_time_to_filetime(props.last_saved_time))
        .add_lpstr(0x12, props.application_name.as_bytes().to_vec())
        .add_i4(0x13, props.security)
        .build(OLE_METADATA_STREAM_SIZE)
}

/// 构造 DocumentSummaryInformation 流（固定 4096 字节）
pub fn build_document_summary_information_stream(props: &DocSummaryInfoProps) -> Vec<u8> {
    // 将 heading_pairs 转为 Variant 序列
    let mut variants = Vec::with_capacity(props.heading_pairs.len() * 2);
    for (name, count) in &props.heading_pairs {
        variants.push(Variant::Lpstr(name.clone()));
        variants.push(Variant::I4(*count));
    }

    SectionBuilder::new(FMTID_DOC_SUMMARY_INFORMATION)
        .add_i2(0x01, props.codepage)
        .add_i4(0x17, props.version)
        .add_bool(0x0B, props.scale_crop)
        .add_bool(0x10, props.links_dirty)
        .add_bool(0x13, props.shared_doc)
        .add_bool(0x16, props.hlinks_changed)
        .add_vector_lpstr(0x0D, props.titles_of_parts.clone())
        .add_vector_variant(0x0C, variants)
        .build(OLE_METADATA_STREAM_SIZE)
}

// ============================================================================
// 测试
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_summary_stream_size() {
        let props = SummaryInfoProps::default();
        let stream = build_summary_information_stream(&props);
        assert_eq!(stream.len(), OLE_METADATA_STREAM_SIZE);
    }

    #[test]
    fn test_doc_summary_stream_size() {
        let props = DocSummaryInfoProps::default();
        let stream = build_document_summary_information_stream(&props);
        assert_eq!(stream.len(), OLE_METADATA_STREAM_SIZE);
    }

    #[test]
    fn test_filetime_conversion() {
        let ft = system_time_to_filetime(UNIX_EPOCH);
        // 1970-01-01 对应的 FILETIME 应为 116444736000000000
        assert_eq!(ft, 116444736000000000);
    }
}
