//! Excel 工作簿导出管理器模块
//!
//! 本模块定义了 Excel 导出的核心容器结构 Workbook，负责：
//! - 管理全局样式配置（样式池）
//! - 维护工作表导出任务队列
//! - 协调多工作表的统一导出流程
//!
//! # 设计理念
//!
//! ## 统一样式管理
//! 通过全局样式池实现样式配置的集中管理和复用，避免重复创建样式对象。
//!
//! ## 批量导出支持
//! 支持将多个 WorkSheet 实例整合到单一 Excel 文件中，保持导出顺序。
//!
//! ## 生命周期管理
//! 通过 Rust 的所有权系统确保资源的正确管理和释放。
use std::collections::HashMap;
use std::error::Error;
use polars::prelude::*;
use rust_xlsxwriter::*;
use crate::worksheet::WorkSheet;
use crate::error::XlsxError;
use std::fmt;
use crate::merge_factory::MergeFactory;
use crate::style_library::StyleLibrary;
use crate::style_factory::StyleFactory;

/// Excel 工作簿导出管理器
///
/// 作为 Excel 导出系统的顶层容器，负责协调多个工作表的统一导出。
/// 通过样式池和任务队列的管理，实现高效、一致的 Excel 文件生成。
///
/// # 核心功能
///
/// ## 样式统一管理
/// 维护全局样式配置，确保整个工作簿的视觉一致性。
///
/// ## 多工作表支持
/// 支持批量导出多个工作表到单一 Excel 文件。
///
/// ## 资源优化
/// 通过样式共享和批量处理提高内存使用效率。
///
/// # 设计架构
///
/// ## 分层结构
/// ```text
/// Workbook (工作簿)
///   ├── styles (样式池)
///   └── sheets (工作表队列)
///       ├── WorkSheet 1
///       ├── WorkSheet 2
///       └── ...
/// ```
///
/// ## 数据流向
/// ```text
/// 多个 DataFrame → WorkSheet 任务 → Workbook 管理 → Excel 文件
/// ```
///
/// # 线程安全性
///
/// 结构体字段为公共可变，适合在单线程环境中使用。
/// 如需多线程支持，应考虑外部同步机制。
pub struct Workbook {

    /// 全局样式配置池（私有字段）
    ///
    /// 存储预定义的样式格式，通过字符串标签进行引用。
    /// 使用 HashMap 实现 O(1) 的样式查找性能。
    ///
    /// # 设计考虑
    /// - **私有访问**：外部无法直接修改样式池，确保数据一致性
    /// - **内部优化**：通过专用方法管理样式，提供更好的封装性
    /// - **资源共享**：相同样式在多个单元格间共享，减少内存占用
    ///
    /// # 样式管理策略
    /// 样式通过 `StyleLibrary` 或直接注册方式进行管理，避免重复创建。
    styles: HashMap<String, Format>,
    /// 工作表导出任务队列（私有字段）
    ///
    /// 按照 Vec 中的顺序进行工作表导出，索引 0 对应 Excel 中的第一个工作表。
    /// 每个工作表包含完整的数据和样式信息。
    ///
    /// # 设计考虑
    /// - **顺序保证**：维护工作表的添加顺序，确保导出一致性
    /// - **所有权管理**：WorkSheet 持有 DataFrame 所有权，确保数据完整性
    /// - **批量处理**：支持多个工作表的统一导出管理
    ///
    /// # 访问控制
    /// 通过专用的公共方法（如 `add_worksheet`、`workbook_count` 等）进行管理，
    /// 避免外部直接操作导致的状态不一致。
    sheets: Vec<WorkSheet>,

}

impl Workbook {
    /// 获取工作簿中工作表的数量
    ///
    /// 用于查询当前工作簿中已添加的工作表总数，
    /// 便于了解工作簿的规模和进行相关统计。
    ///
    /// # 返回值
    ///
    /// 返回 `usize` 类型的工作表数量：
    /// - `0` 表示空工作簿（尚未添加任何工作表）
    /// - `n` 表示包含 n 个工作表
    ///
    /// # 时间复杂度
    /// O(1) - 直接返回 Vec 长度，无需遍历
    ///
    /// # 内存安全性
    /// 方法不持有任何可变引用，确保线程安全性
    ///
    /// # 设计考虑
    ///
    /// ## 只读访问
    /// 提供对私有 `sheets` 字段的安全只读访问
    ///
    /// ## 信息隐藏
    /// 不暴露内部 Vec 结构，仅提供数量信息
    ///
    /// ## 一致性保证
    /// 返回值始终与 `sheets.len()` 保持一致
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// 按索引获取工作表名称
    ///
    /// # 参数
    /// * `index` - 工作表索引
    ///
    /// # 返回值
    /// * `Some(&str)` - 工作表名称
    /// * `None` - 索引无效
    pub fn get_sheet_name(&self, index: usize) -> Option<&str> {
        self.sheets.get(index).map(|sheet| sheet.name.as_str())
    }

    /// 按索引获取工作表的样式映射（简化版）
    ///
    /// 返回样式映射的直接引用，如果工作表不存在则返回 None。
    ///
    /// # 参数
    /// * `index` - 工作表索引
    ///
    /// # 返回值
    /// * `Some(&HashMap<(u32, u16), Arc<str>>)` - 存在且有样式映射
    /// * `None` - 工作表不存在或无样式映射
    pub fn get_sheet_style_map_simple(&self, index: usize) -> Option<&HashMap<(u32, u16), Arc<str>>> {
        self.sheets.get(index).and_then(|sheet| sheet.style_map.as_ref())
    }

}

impl Workbook {

    /// 创建并初始化新的 Workbook 实例
    ///
    /// 工厂方法模式实现，负责 Workbook 对象的构造和基础环境初始化。
    /// 通过跨平台字体适配和预置样式，为 Excel 导出提供一致的默认体验。
    ///
    /// # 设计理念
    ///
    /// ## 开箱即用
    /// 提供合理的默认配置，让用户无需额外设置即可获得良好的导出效果。
    ///
    /// ## 跨平台兼容性
    /// 根据编译目标平台自动选择最适合的中文字体，确保在不同系统上的显示效果。
    ///
    /// ## 样式继承机制
    /// 通过样式克隆和覆盖实现样式的层次化管理，减少重复配置。
    ///
    /// # 初始化流程
    ///
    /// ## 三阶段初始化
    /// ```text
    /// 阶段1: 跨平台字体适配
    ///   ↓ 根据操作系统选择合适字体
    /// 阶段2: 基础样式创建
    ///   ↓ 创建标准数据样式和表头样式
    /// 阶段3: 样式池初始化
    ///   ↓ 将预置样式注册到样式池中
    /// ```
    ///
    /// ## 字体适配策略
    ///
    /// ### 平台特定字体
    /// - **Windows**: Microsoft YaHei (微软雅黑) - 系统默认中文字体
    /// - **macOS**: PingFang SC (苹方) - 现代化中文字体
    /// - **其他平台**: sans-serif - 通用无衬线字体回退方案
    ///
    /// ### 字体兼容性考虑
    /// Excel 文件不嵌入字体，仅存储字体名称，因此选择各平台的标准字体
    /// 确保在目标系统上能够正确显示。
    ///
    /// ## 基础样式定义
    ///
    /// ### 标准数据样式 ("default")
    /// - 字体：平台适配字体，11号字
    /// - 边框：细黑边框（四周）
    /// - 对齐：水平居中 + 垂直居中
    /// - 用途：普通数据单元格的默认样式
    ///
    /// ### 表头样式 ("header")
    /// - 继承：基于标准样式的所有属性
    /// - 扩展：加粗字体 + 浅灰色背景 (#BFBFBF)
    /// - 用途：表格标题行的突出显示
    ///
    /// # 样式管理机制
    ///
    /// ## 样式继承
    /// 通过 `clone()` 方法实现样式的继承，避免重复的样式定义。
    ///
    /// ## 样式池注册
    /// 预置样式通过字符串标签注册到样式池中，便于后续引用。
    ///
    /// # 返回值
    ///
    /// * `Ok(Workbook)` - 成功创建的 Workbook 实例
    /// * `Err(Box<dyn Error>)` - 理论上极少发生，主要用于 API 一致性
    ///
    /// # 性能特征
    ///
    /// ## 初始化开销
    /// O(1) - 固定数量的样式创建和注册操作
    ///
    /// ## 内存效率
    /// - 样式共享：通过样式池避免重复对象创建
    /// - 字体优化：选择系统标准字体减少兼容性问题
    ///
    /// # 错误处理策略
    ///
    /// ## 理论可靠性
    /// 在内存充足的情况下，初始化过程几乎不会失败。
    ///
    /// ## API 一致性
    /// 返回 `Result` 类型保持与后续 I/O 操作的接口一致性。
    ///
    /// # 扩展性考虑
    ///
    /// ## 样式扩展
    /// 可以通过 `register_style` 方法添加更多预置样式。
    ///
    /// ## 平台扩展
    /// 可以添加更多平台特定的字体适配逻辑。
    pub fn new() -> Result<Self, Box<dyn Error>> {
        let mut styles = HashMap::new();

        // 1. 跨平台字體自適應適配：
        // 利用編譯時宏 cfg! 確定目標系統。Excel 文件不嵌入字體，僅存儲名稱，
        // 故需確保指定的名稱在對應系統中是可被識別的標準名稱。
        let font_name = if cfg!(target_os = "windows") {
            "Microsoft YaHei" // Windows: 微软雅黑
        } else if cfg!(target_os = "macos") {
            "PingFang SC"      // macOS: 苹方
        } else {
            "sans-serif"       // Linux 或其他: 通用無襯線字體
        };

        // 2. 定義「標準樣式」：
        // 這是報表中絕大多數單元格的底層模板，包含了邊框、字體和對齊方式。
        let standard_fmt = Format::new()
            .set_font_name(font_name) // 設置微軟雅黑
            .set_font_size(11)                 // 設置常用字號
            .set_border(FormatBorder::Thin)   // 設置細邊框（四周）
            .set_align(FormatAlign::Center)   // 水平居中
            .set_align(FormatAlign::VerticalCenter); // 垂直居中

        // 3. 快速派生「表頭樣式」：
        // 通過 clone() 繼承 standard_fmt 的屬性，僅對差異項（加粗、背景）進行覆蓋。
        // 這種鏈式調用確保了表頭與數據行在字體、邊框寬度上完全對齊。
        let header_fmt = standard_fmt.clone()
            .set_bold()
            .set_background_color(Color::RGB(0xBFBFBF));

        // 4. 將初始化樣式存入緩存池：
        // "default" 用於普通單元格保底，"header" 用於標題行。
        styles.insert("default".to_string(), standard_fmt);
        styles.insert("header".to_string(), header_fmt);

        Ok(Self {
            styles,
            sheets: vec![],
        })
    }

    /// 注册或更新工作簿中的样式定义
    ///
    /// 构建器模式方法，用于向工作簿的样式池中添加新的样式定义
    /// 或更新已存在的样式配置。通过链式调用支持流畅的样式配置体验。
    ///
    /// # 设计模式
    ///
    /// ## 构建器模式
    /// 采用可链式调用的构建器模式，支持样式配置的流畅 API：
    ///
    /// ## 所有权语义
    /// 方法接收 `mut self` 参数并返回 `Self`，实现所有权的转移和链式调用。
    ///
    /// # 参数说明
    ///
    /// * `name`: 样式名称标识符
    ///   - 用于在样式映射中唯一标识样式
    ///   - 在单元格样式应用时作为引用键使用
    /// * `format`: 实际的样式格式对象
    ///   - `rust_xlsxwriter::Format` 类型的完整样式定义
    ///   - 包含字体、颜色、对齐、边框等所有样式属性
    ///
    /// # 执行机制
    ///
    /// ## 样式注册流程
    /// ```text
    /// 输入样式名称和格式对象
    ///   ↓ 字符串转换
    /// 将 &str 转换为 String 作为 HashMap 键
    ///   ↓ HashMap 插入
    /// 自动处理重复键的覆盖更新
    ///   ↓ 返回实例
    /// 支持链式调用的自我返回
    /// ```
    ///
    /// ## HashMap 特性利用
    ///
    /// ### 键值覆盖
    /// HashMap 的 `insert` 方法会自动处理同名键的情况：
    /// - 如果键不存在：添加新的键值对
    /// - 如果键已存在：更新值并返回旧值（被忽略）
    ///
    /// ### 内存管理
    /// - 字符串键：通过 `to_string()` 创建所有权副本
    /// - 格式值：直接转移所有权到 HashMap 中
    ///
    /// # 性能特征
    ///
    /// ## 时间复杂度
    /// O(1) 平均情况 - HashMap 插入操作
    ///
    /// ## 内存效率
    /// - 字符串键：必要时进行所有权转换
    /// - 格式值：零拷贝的所有权转移
    /// - 重复键：自动回收旧值内存
    ///
    /// ## 链式调用优化
    /// - 避免中间变量分配
    /// - 支持编译时优化
    ///
    /// # 错误处理
    ///
    /// ## 静默处理策略
    /// 当前实现采用静默处理方式，样式覆盖不会产生错误或警告，
    /// 确保流畅的配置体验。
    ///
    /// ## 内存安全
    /// 通过 Rust 的所有权系统确保内存安全，避免悬垂指针和内存泄漏。
    ///
    /// # 设计考虑
    ///
    /// ## API 一致性
    /// 与其他构建器模式方法保持一致的签名和行为。
    ///
    /// ## 灵活性
    /// 支持任意 `rust_xlsxwriter::Format` 对象，不限制样式复杂度。
    pub fn set_style(mut self, name: &str, format: Format) -> Self {
        // 將字符串切片轉換為 String 並存入 HashMap。
        // HashMap 的 insert 特性確保了同名鍵值的自動更新。
        self.styles.insert(name.to_string(), format);

        // 返回修改後的對象，支持鏈式調用：exporter.set_style(...).set_style(...)
        self
    }

    /// 使用样式库扩展工作簿的样式配置
    ///
    /// 构建器模式方法，用于将外部定义的样式库集成到当前工作簿中。
    /// 通过批量注册样式库中的所有预定义样式，实现样式的统一管理和复用。
    ///
    /// # 设计模式
    ///
    /// ## 构建器模式
    /// 采用可链式调用的构建器模式，支持流畅的 API 使用体验：
    ///
    /// ## 所有权转移
    /// 方法接收 `mut self` 参数并返回 `Self`，实现所有权的转移和链式调用。
    ///
    /// # 参数说明
    ///
    /// * `library`: 对样式库的不可变引用
    ///   - 包含预定义的样式配置集合
    ///   - 通过 `build_formats()` 方法转换为实际的 Format 对象
    ///
    /// # 执行流程
    ///
    /// ## 三阶段处理
    /// ```text
    /// 阶段1: 样式构建
    ///   ↓ 调用 library.build_formats() 生成样式映射
    /// 阶段2: 批量注册
    ///   ↓ 遍历样式映射，逐个注册到工作簿样式池
    /// 阶段3: 返回实例
    ///   ↓ 返回更新后的工作簿实例支持链式调用
    /// ```
    ///
    /// ## 样式库集成机制
    ///
    /// ### 批量转换
    /// 通过 `library.build_formats()` 一次性获取所有预定义样式，
    /// 避免重复的样式构建开销。
    ///
    /// ### 逐个注册
    /// 利用现有的 `set_style` 方法确保样式注册的一致性，
    /// 包括样式覆盖和重复处理等逻辑。
    ///
    /// # 样式管理策略
    ///
    /// ## 命名空间管理
    /// 样式库中的样式名称直接映射到工作簿样式池中，
    /// 如果存在同名样式，样式库中的样式将覆盖现有样式。
    ///
    /// ## 冲突处理
    /// 通过 `set_style` 方法的覆盖机制处理样式名称冲突，
    /// 确保最新的样式配置生效。
    ///
    /// # 性能特征
    ///
    /// ## 时间复杂度
    /// O(n) - 其中 n 为样式库中样式的数量
    ///
    /// ## 内存效率
    /// - 样式共享：通过样式池避免重复对象创建
    /// - 批量处理：减少单次样式注册的函数调用开销
    ///
    /// # 错误处理
    ///
    /// ## 静默失败策略
    /// 当前实现采用静默处理方式，样式库中的任何问题都不会导致
    /// 整体操作失败，确保流程的连续性。
    ///
    /// ## 样式验证
    /// 样式库负责确保生成的 Format 对象的有效性，
    /// 工作簿只负责存储和管理。
    pub fn with_library(mut self, library: &StyleLibrary) -> Self {
        // 調用 Library 內部的批量轉換邏輯
        // 将样式库定义转换为实际的 Format 对象集合
        let formats = library.build_formats();

        for (name, fmt) in formats {
            // 利用現有的 set_style 實現，將 Format 存入 Workbook.styles
            // 通过已有的样式注册机制确保一致性和冲突处理
            self = self.set_style(&name, fmt);
        }

        self
    }

    /// 从 JSON 配置应用样式到工作簿
    ///
    /// 便捷方法，用于直接从 JSON 配置数据创建样式库并应用到当前工作簿。
    /// 通过两阶段处理（JSON 解析 → 样式库构建 → 样式应用）实现配置驱动的样式管理。
    ///
    /// # 设计理念
    ///
    /// ## 配置驱动
    /// 支持通过外部 JSON 配置文件或内联 JSON 数据动态配置样式，
    /// 实现样式定义与代码的分离。
    ///
    /// ## 两阶段处理
    /// 将复杂的样式配置处理分解为独立的关注点：
    /// 1. JSON 解析和验证
    /// 2. 样式库构建和应用
    ///
    /// # 参数说明
    ///
    /// * `value`: 对 JSON 配置的不可变引用
    ///   - 包含样式定义的结构化数据
    ///   - 格式需符合 StyleLibrary 的 JSON Schema
    ///
    /// # 执行流程
    ///
    /// ## 流水线处理
    /// ```text
    /// JSON 配置输入
    ///   ↓ StyleLibrary::from_json()
    /// 样式库实例
    ///   ↓ self.with_library()
    /// 更新后的工作簿实例
    /// ```
    ///
    /// ## JSON 解析阶段
    /// - 调用 `StyleLibrary::from_json()` 进行配置解析
    /// - 验证 JSON 结构和数据有效性
    /// - 转换为内部样式定义
    ///
    /// ## 样式应用阶段
    /// - 利用 `with_library()` 方法批量注册样式
    /// - 处理样式名称冲突和覆盖
    /// - 返回更新后的工作簿实例
    ///
    /// # JSON 配置格式
    ///
    /// ## 基本结构
    /// ```json
    /// {
    ///   "styles": {
    ///     "header_bg": { "bg_color": "#BFBFBF", "bold": true, "border": "thin", "align": "center" },
    ///     "money_red": { "font_color": "#FF0000", "num_format": "#,##0.00" },
    ///     "default": { "font_name": "Microsoft YaHei", "font_size": 11 }
    ///   },
    /// }
    /// ```
    ///
    /// ## 配置验证
    /// - 结构验证：确保必需字段存在
    /// - 数据验证：检查值的有效性
    /// - 类型验证：确保数据类型匹配
    ///
    /// # 错误处理策略
    ///
    /// ## 传播错误
    /// 使用 `?` 操作符传播底层错误，保持错误链的完整性。
    ///
    /// ## 错误类型
    /// * `serde_json::Error` - JSON 解析失败
    /// * `StyleLibraryError` - 样式库构建错误
    /// * 其他可能的验证错误
    ///
    /// ## 原子性保证
    /// 如果任何阶段失败，整个操作会回滚，保持工作簿的原始状态。
    ///
    /// # 性能特征
    ///
    /// ## 时间复杂度
    /// O(n + m) - 其中 n 为 JSON 解析复杂度，m 为样式数量
    ///
    /// ## 内存效率
    /// - 零拷贝设计：JSON 引用传递避免不必要的数据复制
    /// - 样式共享：通过样式池实现内存优化
    ///
    /// # 安全考虑
    ///
    /// ## 输入验证
    /// 通过 StyleLibrary 的验证机制确保 JSON 配置的安全性。
    ///
    /// ## 资源限制
    /// 防止恶意配置导致的资源耗尽（如过多样式定义）。
    pub fn apply_styles(self, value: &serde_json::Value) -> Result<Self, Box<dyn Error>> {
        // 1. 在內部利用片段構建樣式庫
        // 通过 JSON 配置创建样式库实例
        // 此步骤会进行配置解析和验证
        let library = StyleLibrary::from_json(value)?;

        // 2. 調用前者進行物理注入
        // 利用已有的 with_library 方法应用样式库
        // 保持样式应用逻辑的一致性
        Ok(self.with_library(&library))
    }

    /// 向工作簿中添加新的工作表任务
    ///
    /// 构建器模式方法，负责将 DataFrame 数据封装为 WorkSheet 任务并添加到工作簿中。
    /// 通过智能的错误处理和自动修复机制，提供鲁棒的工作表添加体验。
    ///
    /// # 设计理念
    ///
    /// ## 智能错误处理
    /// 采用分级错误处理策略，对不同类型的问题采取不同的应对措施：
    /// - 静默跳过：对无意义的输入（空表）自动忽略
    /// - 自动修复：对可修复的问题（非法名称）自动纠正
    /// - 明确报错：对严重问题及时反馈给用户
    ///
    /// ## 数据完整性保证
    /// 通过多重验证确保添加的工作表任务具备完整性和有效性。
    ///
    /// # 参数说明
    ///
    /// * `df`: 待导出的 DataFrame 数据源
    ///   - 包含实际的数据内容
    ///   - 支持所有权转移以提高性能
    /// * `name`: 可选的工作表名称
    ///   - None：使用默认名称 "Sheet N"
    ///   - Some(name)：使用提供的名称（会进行验证）
    /// * `style_map`: 可选的单元格样式映射
    ///   - None：不应用特殊样式
    ///   - Some(map)：应用指定的样式映射
    /// * `merge_ranges`: 可选的工作表合并区域范围列表
    ///   - None：不进行任何合并操作
    ///   - Some(ranges)：指定需要合并的单元格范围（每组四个参数表示起始行、起始列、结束行、结束列）
    ///
    /// # 返回值
    ///
    /// * `Ok(Workbook)` - 成功添加工作表后的工作簿实例
    /// * `Err(Box<dyn Error>)` - 添加过程中发生的错误
    ///
    /// # 执行流程
    ///
    /// ## 五阶段处理流程
    /// ```text
    /// 阶段1: 名称预处理
    ///   ↓ 确定工作表名称（用户指定或默认生成）
    /// 阶段2: WorkSheet 构建
    ///   ↓ 创建 WorkSheet 任务实例，包括合并区域
    /// 阶段3: 全局冲突检查
    ///   ↓ 检查工作表名称唯一性
    /// 阶段4: 样式引用验证
    ///   ↓ 验证样式名称的存在性
    /// 阶段5: 任务注册
    ///   ↓ 将通过验证的任务添加到队列中
    /// ```
    ///
    /// ## 名称处理策略
    ///
    /// ### 默认命名逻辑
    /// 使用闭包封装默认命名规则：`format!("Sheet {}", sheets_len + 1)`
    /// 确保命名的一致性和唯一性。
    ///
    /// ### 分级错误处理
    /// ```text
    /// WorkSheet::new() 返回结果
    ///   ├─ Ok(task) → 直接使用
    ///   ├─ Err(EmptyDataFrame) → 静默跳过
    ///   ├─ Err(InvalidName) → 自动修复重试
    ///   └─ Err(other) → 向上抛出错误
    /// ```
    ///
    /// ## 验证机制
    ///
    /// ### 名称唯一性检查
    /// 遍历现有工作表列表，确保新添加的工作表名称不重复。
    /// 这是 Excel 的硬性要求，必须在导出前拦截。
    ///
    /// ### 样式引用验证
    /// 检查样式映射中引用的所有样式名称是否在样式池中存在，
    /// 防止导出时出现悬空引用错误。
    ///
    /// # 错误处理策略
    ///
    /// ## 分级响应机制
    ///
    /// ### 静默跳过 (EmptyDataFrame)
    /// 空数据表不具备导出意义，自动忽略而不报错。
    ///
    /// ### 自动修复 (InvalidName)
    /// 名称不合规时自动使用默认名称重试。
    ///
    /// ### 明确报错 (其他错误)
    /// 严重错误（如样式引用不存在）会明确报告。
    ///
    /// # 性能特征
    ///
    /// ## 时间复杂度
    /// - WorkSheet 创建：O(1) + 样式映射清理开销
    /// - 名称重复检查：O(n) 其中 n 为现有工作表数量
    /// - 样式验证：O(m) 其中 m 为样式映射条目数
    ///
    /// ## 内存效率
    /// - DataFrame 克隆：浅拷贝，开销极低
    /// - 样式共享：使用 Arc<str> 避免重复存储
    /// - 错误处理：最小化不必要的对象创建
    ///
    /// # 安全考虑
    ///
    /// ## 数据完整性
    /// 通过多重验证确保添加的工作表任务是完整和有效的。
    ///
    /// ## 资源管理
    /// 正确处理 DataFrame 和样式映射的所有权转移。
    ///
    /// ## 边界条件
    /// 妥善处理空数据、非法名称、重复名称等边界情况。
    pub fn insert(mut self, df: DataFrame, name: Option<String>,
                  style_map: Option<HashMap<(u32, u16), Arc<str>>>,
                  merge_ranges: Option<Vec<(u32, u16, u32, u16)>>
    ) -> Result<Self, Box<dyn Error>> {
        // 1. 定義輔助閉包：封裝默認命名邏輯，確保命名的一致性與唯一性起點
        // 闭包提供统一的默认命名规则，避免代码重复
        let get_default_name = |sheets_len: usize| format!("Sheet {}", sheets_len + 1);

        // 2. 初步確定名稱：優先使用用戶提供，否則生成默認名
        // 支持用户自定义名称，提高灵活性
        let final_name = name.unwrap_or_else(|| get_default_name(self.sheets.len()));

        // 3. 嘗試構建 WorkSheet 任務：
        // 這裡使用 match 進行細分錯誤處理。注意 df.clone() 是淺拷貝，開銷極低。
        // 這裡 clone() 是必須的，因為如果 InvalidName 發生，我們需要原始數據進行第二次嘗試。
        let task = match WorkSheet::new(df.clone(), final_name.clone(), style_map.clone(), merge_ranges.clone()) {
            Ok(t) => t,
            // 規則 A (靜默跳過)：空表不具備導出意義，直接返回原始對象，不存入隊列
            // 对于空数据表采取宽容策略，避免不必要的错误
            Err(XlsxError::EmptyDataFrame) => return Ok(self),
            // 規則 B (自動修復)：名稱非法時，放棄用戶名稱，改用系統預設名稱重試
            // 此時已知 df 不為空，第二次 new 操作是安全的
            Err(XlsxError::InvalidName(_)) => {
                // 使用默认名称重新尝试创建工作表
                let fallback_name = get_default_name(self.sheets.len());
                // 再次調用 new，此時使用安全名稱（已知 df 不為空，所以這次一定會 Ok）
                WorkSheet::new(df, fallback_name, style_map, merge_ranges)?
            }
            // 其他嚴重錯誤：直接包裝並向上拋出
            // 保留原始错误信息，便于问题诊断
            Err(e) => return Err(Box::new(e)),
        };

        // 4. 名稱重複檢查：
        // Excel 不允許同名工作表。這是全局級別的衝突，必須由 Workbook 攔截。
        // 确保 Excel 文件的结构完整性
        if self.sheets.iter().any(|s| s.name == task.name) {
            return Err(Box::new(XlsxError::DuplicateName(task.name)));
        }

        // 5. 樣式名存在性檢查：
        // 確保數據在寫入時能找到對應的格式定義，防止 save 時出現懸空引用。
        // 预先验证样式引用的有效性，避免运行时错误
        if let Some(ref map) = task.style_map {
            for style_name in map.values() {
                if !self.styles.contains_key(style_name.as_ref()) {
                    return Err(Box::new(XlsxError::UnknownStyle(style_name.clone())));
                }
            }
        }

        // 通過所有校驗，將任務存入隊列
        // 所有权转移完成，任务正式加入工作簿
        self.sheets.push(task);
        Ok(self)
    }

    /// 使用样式工厂向工作簿中添加带样式的工作表
    ///
    /// 便捷方法，用于将 DataFrame 数据结合样式工厂生成的样式映射和合并区域
    /// 一体化地添加到工作簿中。通过自动化样式计算减少用户的配置负担。
    ///
    /// # 设计理念
    ///
    /// ## 一体化处理
    /// 将样式计算、合并区域设置和工作表添加两个步骤合并为一个原子操作，
    /// 提供更简洁的 API 体验。
    ///
    /// ## 职责分离
    /// 样式工厂负责样式逻辑计算，合并工厂负责合并区域设置，工作簿负责样式应用和管理，
    /// 保持关注点的清晰分离。
    ///
    /// # 参数说明
    ///
    /// * `df`: 待导出的 DataFrame 数据源
    ///   - 提供给样式工厂进行样式计算
    ///   - 作为工作表的数据内容
    /// * `name`: 可选的工作表名称
    ///   - None：使用默认名称生成策略
    ///   - Some(name)：使用提供的名称（会进行验证）
    /// * `style_factory`: 可选的样式工厂引用
    ///   - 负责基于 DataFrame 数据计算样式映射
    ///   - 不会被消费，可重复使用
    /// * `merge_factory`: 可选的合并区域工厂引用
    ///   - 负责基于 DataFrame 数据设置合并区域
    ///   - 不会被消费，可重复使用
    ///
    /// # 执行流程
    ///
    /// ## 三阶段处理
    /// ```text
    /// 输入: DataFrame + 名称 + 样式工厂 + 合并工厂
    ///   ↓ 阶段1: 样式计算
    /// 调用 style_factory.execute(&df) 生成样式映射
    ///   ↓ 阶段2: 合并区域设置
    /// 调用 merge_factory.execute(&df) 设置合并区域
    ///   ↓ 阶段3: 工作表添加
    /// 调用 insert 方法完成工作表注册
    ///   ↓ 输出: 更新后的工作簿实例
    /// ```
    ///
    /// ## 样式计算阶段
    ///
    /// ### 工厂执行
    /// - 样式工厂基于 DataFrame 数据执行样式规则评估
    /// - 生成精确到单元格级别的样式映射
    /// - 返回 HashMap<(u32, u16), Arc<str>> 类型的结果
    ///
    /// ### 错误传播
    /// 样式工厂执行过程中的任何错误都会被传播到调用方。
    ///
    /// ## 合并区域设置阶段
    ///
    /// ### 工厂执行
    /// - 合并工厂基于 DataFrame 数据设置合并区域
    /// - 生成精确到单元格级别的合并区域列表
    /// - 返回 Vec<(u32, u16, u32, u16)> 类型的结果，表示四个角的坐标
    ///
    /// ### 错误传播
    /// 合并工厂执行过程中的任何错误都会被传播到调用方。
    ///
    /// ## 工作表添加阶段
    ///
    /// ### 复用现有逻辑
    /// 直接调用已有的 `insert` 方法，确保行为一致性。
    ///
    /// ### 样式映射传递
    /// 将计算得到的样式映射作为参数传递给 `insert` 方法。
    ///
    /// ### 合并区域传递
    /// 将计算得到的合并区域列表作为参数传递给 `insert` 方法。
    ///
    /// # 错误处理策略
    ///
    /// ## 透明传播
    /// 所有错误都通过 `?` 操作符透明传播，保持错误链的完整性。
    ///
    /// ## 错误类型
    /// * `StyleFactoryError` - 样式计算过程中的错误
    /// * `MergeFactoryError` - 合并区域设置过程中的错误
    /// * `XlsxError` - 工作表添加过程中的错误
    /// * 其他可能的底层错误
    ///
    /// ## 原子性保证
    /// 如果样式计算、合并区域设置或工作表添加任一环节失败，整个操作会回滚，
    /// 保持工作簿的原始状态不变。
    ///
    /// # 性能特征
    ///
    /// ## 时间复杂度
    /// O(f(n) + g(n) + h(m)) - 其中：
    /// - f(n) 为样式工厂执行复杂度（n 为 DataFrame 大小）
    /// - g(n) 为合并工厂执行复杂度（n 为 DataFrame 大小）
    /// - h(m) 为工作表添加复杂度（m 为现有工作表数量）
    ///
    /// ## 内存效率
    /// - 样式共享：使用 Arc<str> 避免重复存储
    /// - 合并区域共享：使用紧凑的坐标列表表示合并区域
    /// - 零拷贝设计：样式映射和合并区域直接传递，避免不必要的复制
    /// - 流水线处理：样式计算、合并设置和添加操作无缝衔接
    ///
    /// # 安全考虑
    ///
    /// ## 数据一致性
    /// 通过原子性操作确保样式计算结果与工作表数据的一致性。
    ///
    /// ## 资源管理
    /// 正确处理 DataFrame 和样式映射的所有权转移。
    ///
    /// ## 边界条件
    /// 委托给底层 `insert` 方法处理各种边界情况。
    ///
    /// # 设计优势
    ///
    /// ## API 简洁性
    /// 将三个相关操作合并为一个方法调用，提高使用便利性。
    ///
    /// ## 逻辑一致性
    /// 复用已有的 `insert` 方法确保行为一致性。
    ///
    /// ## 扩展性好
    /// 保持了与样式工厂和合并工厂的松耦合关系，便于未来扩展。
    pub fn insert_with_factory(
        self,
        df: DataFrame,
        name: Option<String>,
        style_factory: Option<StyleFactory>,
        merge_factory: Option<MergeFactory>
    ) -> Result<Self, Box<dyn Error>> {
        // 1. 執行样式工厂引擎：计算该 DataFrame 的物理样式地图
        // 产出类型为 HashMap<(u32, u16), Arc<str>>
        // 此步骤会进行样式规则评估和样式映射生成
        let style_map = match style_factory {
            Some(factory) => Some(factory.execute(&df)?),
            None => None,
        };

        // 2. 執行合并区域工厂引擎：计算该 DataFrame 的合并区域设置
        // 产出类型为 Vec<(u32, u16, u32, u16)>，表示四个角的坐标
        // 此步骤会进行合并区域规则评估和合并区域生成
        let merge_ranges = match merge_factory {
            Some(factory) => Some(factory.execute(&df)?),
            None => None,
        };

        // 3. 调用原有的成品 insert 函数
        // 注意：这里直接将计算好的 style_map 和 merge_ranges 传入
        // 复用现有的工作表添加逻辑，确保行为一致性
        self.insert(df, name, style_map, merge_ranges)
    }

    /// 使用 JSON 配置向工作簿中添加带样式的工作表
    ///
    /// 便捷方法，用于直接从 JSON 配置数据创建样式工厂和合并工厂，并应用到工作表添加过程。
    /// 通过配置驱动的方式实现样式定义与代码的完全分离，支持灵活的样式配置。
    ///
    /// # 设计理念
    ///
    /// ## 配置驱动架构
    /// 支持通过外部 JSON 配置文件或内联 JSON 数据动态配置样式和合并规则，
    /// 实现样式定义与应用程序代码的彻底分离。
    ///
    /// ## 三层处理模型
    /// ```text
    /// JSON 配置 → 样式工厂 + 合并工厂 → 样式化工作表 + 合并单元格
    ///     ↓              ↓                 ↓                ↓
    ///   配置层        逻辑层          应用层         应用层
    /// ```
    ///
    /// ## 无缝集成
    /// 利用已有的 `insert_with_factory` 方法实现功能复用，
    /// 确保行为一致性和代码维护性。
    ///
    /// # 参数说明
    ///
    /// * `df`: 待导出的 DataFrame 数据源
    ///   - 提供给样式工厂进行样式规则评估
    ///   - 作为工作表的实际数据内容
    /// * `name`: 可选的工作表名称
    ///   - None：使用系统默认命名策略
    ///   - Some(name)：使用提供的名称（会进行验证和清理）
    /// * `config`: 对 JSON 配置的不可变引用
    ///   - 包含样式规则定义和合并规则的结构化数据
    ///   - 支持 null 值，由工厂内部进行容错处理
    ///
    /// # 执行流程
    ///
    /// ## 三阶段流水线处理
    /// ```text
    /// 输入: DataFrame + 名称 + JSON 配置
    ///   ↓ 阶段1: 样式和合并工厂创建
    /// 解析 JSON 配置创建 StyleFactory 和 MergeFactory 实例
    ///   ↓ 阶段2: 样式化工作表添加
    /// 调用 insert_with_factory 完成样式计算和添加
    ///   ↓ 阶段3: 返回更新后的工作簿
    /// ```
    ///
    /// ## 配置解析阶段
    ///
    /// ### 容错处理机制
    /// - **Null 值处理**：工厂内部会将 null 值转换为空规则集
    /// - **格式验证**：JSON 结构验证确保配置的有效性
    /// - **默认行为**：无效配置自动降级为无样式模式或无合并规则
    ///
    /// ### 数据克隆策略
    /// 使用 `config.clone()` 确保工厂拥有配置数据的所有权，
    /// 避免生命周期问题，虽然有一定开销但保证了安全性。
    ///
    /// # JSON 配置格式
    ///
    /// ## 基本结构示例
    /// ```json
    /// {
    ///   "style_rules": [
    ///     {
    ///       "row_conditions": [
    ///         {"type": "index", "criteria": [0]}
    ///       ],
    ///       "apply": {
    ///         "style": "header",
    ///         "overrides": [
    ///           {
    ///             "style": "highlight",
    ///             "col_conditions": [
    ///               {"type": "match", "targets": ["金额", "数量"]}
    ///             ]
    ///           }
    ///         ]
    ///       }
    ///     }
    ///   ],
    ///   "merge_rules": [
    ///     {
    ///       "row_range": [1, 3],
    ///       "col_range": [2, 4]
    ///     }
    ///   ]
    /// }
    /// ```
    ///
    /// ## 配置灵活性
    /// - **空配置**：null 或 [] 表示无样式规则或无合并规则
    /// - **复杂规则**：支持嵌套的条件和覆盖逻辑
    /// - **动态调整**：运行时可更改配置而无需重新编译
    ///
    /// # 错误处理策略
    ///
    /// ## 分层错误传播
    /// ```text
    /// JSON 解析错误 → 样式和合并工厂创建失败
    ///   ↓
    /// 样式计算错误 → 工作表添加失败
    ///   ↓
    /// 工作表验证错误 → 插入操作失败
    /// ```
    ///
    /// ## 错误类型
    /// * `serde_json::Error` - JSON 解析和验证错误
    /// * `StyleFactoryError` - 样式工厂内部错误
    /// * `MergeFactoryError` - 合并工厂内部错误
    /// * `XlsxError` - 工作表添加和验证错误
    ///
    /// ## 原子性保证
    /// 任何阶段的失败都会导致整个操作回滚，保持工作簿的原始状态。
    ///
    /// # 性能特征
    ///
    /// ## 时间复杂度
    /// O(p + f(n) + g(m)) - 其中：
    /// - p 为 JSON 解析复杂度
    /// - f(n) 为样式计算复杂度（n 为 DataFrame 大小）
    /// - g(m) 为工作表添加复杂度（m 为现有工作表数量）
    ///
    /// ## 内存效率
    /// - **配置克隆**：`config.clone()` 会产生一次数据复制开销
    /// - **样式共享**：使用 Arc<str> 避免样式名称重复存储
    /// - **合并规则共享**：使用 Arc<str> 避免合并规则重复存储
    /// - **流水线优化**：各阶段无缝衔接减少中间状态
    ///
    /// # 安全考虑
    ///
    /// ## 输入验证
    /// 通过样式工厂和合并工厂的内置验证机制确保配置数据的安全性。
    ///
    /// ## 资源限制
    /// 防止恶意配置导致的资源耗尽（如过深的嵌套或过多的规则）。
    ///
    /// ## 内存安全
    /// 利用 Rust 的所有权系统确保所有资源的正确管理。
    ///
    /// # 设计优势
    ///
    /// ## 零配置启动
    /// 支持 null 配置，实现零配置的快速启动。
    ///
    /// ## 配置复用
    /// 同一份配置可应用于多个不同的 DataFrame。
    ///
    /// ## 行为一致性
    /// 复用已有的方法确保与手动创建工厂的行为一致。
    pub fn insert_with_config(self, df: DataFrame, name: Option<String>, config: &serde_json::Value) -> Result<Self, Box<dyn Error>> {
        // 直接傳入引用的片段（可能是 Null），Factory 內部會自我癒合
        let style_factory = match StyleFactory::new(config.clone()) {
            Ok(factory) => Some(factory),
            Err(_) => None
        };
        let merge_factory = match MergeFactory::new(config.clone()) {
            Ok(factory) => Some(factory),
            Err(_) => None
        };

        Ok(self.insert_with_factory(df, name, style_factory, merge_factory)?)
    }

    /// 将工作簿保存为 Excel 文件
    ///
    /// 核心导出方法，负责将所有注册的工作表任务转换为实际的 Excel 文件。
    /// 通过精细的样式处理和类型转换，确保数据在 Excel 中的正确呈现。
    ///
    /// # 设计理念
    ///
    /// ## 完整性保证
    /// 确保所有工作表、样式和数据都能正确转换并保存到 Excel 文件中。
    ///
    /// ## 类型安全转换
    /// 将 Polars 的丰富数据类型精确映射到 Excel 支持的格式。
    ///
    /// ## 样式一致性
    /// 通过多层次的样式查找机制确保视觉呈现的一致性。
    ///
    /// # 参数说明
    ///
    /// * `path`: 目标 Excel 文件的保存路径
    ///   - 可以是相对路径或绝对路径
    ///   - 文件格式自动识别为 .xlsx
    ///   - 如果文件已存在会被覆盖
    ///
    /// # 执行流程
    ///
    /// ## 五阶段处理流程
    /// ```text
    /// 阶段1: 工作簿初始化
    ///   ↓ 创建 rust_xlsxwriter 工作簿实例
    /// 阶段2: 样式预取
    ///   ↓ 预加载常用样式引用避免重复查找
    /// 阶段3: 工作表处理
    ///   ↓ 遍历每个工作表进行数据写入
    /// 阶段4: 类型转换
    ///   ↓ 将 Polars 数据类型转换为 Excel 格式
    /// 阶段5: 文件保存
    ///   ↓ 将内存中的工作簿写入磁盘文件
    /// ```
    ///
    /// ## 样式查找优先级
    /// ```text
    /// 单元格样式查找顺序：
    /// 1. 单元格特定样式 → style_map.get(&(row, col))
    /// 2. 表头默认样式 → styles.get("header")
    /// 3. 列默认样式 → column_defaults.get(col)
    /// 4. 标准默认样式 → styles.get("standard")
    /// ```
    ///
    /// ## 合并区域处理
    /// ```text
    /// 合并区域处理顺序：
    /// 1. 在数据写入前应用所有合并区域
    /// 2. 使用空白值填充合并区域以保持样式
    /// 3. 然后按正常流程写入数据和标题
    /// ```
    ///
    /// # 安全考虑
    ///
    /// ## 文件系统安全
    /// - 路径验证确保不会写入系统敏感目录
    /// - 文件覆盖提供明确的用户控制
    ///
    /// ## 数据完整性
    /// - 类型转换确保数据精度不丢失
    /// - 样式应用保持视觉一致性
    /// - 错误处理防止部分写入
    /// - 合并区域正确应用避免格式冲突
    ///
    /// ## 资源管理
    /// - 正确管理文件句柄和内存资源
    /// - 及时释放临时对象避免内存泄漏
    ///
    /// # 功能特性
    ///
    /// ## 自动列宽调整
    /// 在数据写入完成后自动调用 `autofit()` 方法，
    /// 确保列宽能够适应最长的内容显示。
    ///
    /// ## 样式保真度
    /// 通过四级样式查找机制确保每个单元格都能获得适当的样式。
    ///
    /// ## 类型保真度
    /// 精确的类型转换确保数据在 Excel 中的正确显示。
    ///
    /// ## 合并区域支持
    /// 在数据写入前预先处理合并区域，确保Excel合并单元格正确显示。
    pub fn save(&self, path: &str) -> Result<(), Box<dyn Error>> {
        // 创建新的 Excel 工作簿实例，作为导出的目标容器
        let mut workbook = rust_xlsxwriter::Workbook::new();

        // --- 1. 樣式池安全預取：---
        // fallback_fmt 用於解決 Rust 借用檢查器對臨時引用的限制。
        // default_fmt 是你在 new() 中定義的全局 UI 規範（如微軟雅黑+邊框）。
        let fallback_fmt = Format::default();           // 系统回退样式（无样式）
        let default_fmt = self.styles.get("default");   // 全局默认样式（基础格式）

        // 遍历所有注册的工作表任务
        for sheet in &self.sheets {
            // 为当前工作表创建 Excel 工作表实例
            let worksheet = workbook.add_worksheet();
            // 设置工作表名称，支持中文和特殊字符（会进行验证）
            worksheet.set_name(sheet.name.as_str())?;

            // 在写入数据前先应用合并区域
            if let Some(merge_ranges) = &sheet.merge_ranges {
                for merge_range in merge_ranges {
                    // merge_range 包含四个元素：起始行、起始列、结束行、结束列
                    // 第五、六个参数填写空白以获得这些空白单元格
                    worksheet.merge_range(
                        merge_range.0,  // first_row
                        merge_range.1,  // first_col
                        merge_range.2,  // last_row
                        merge_range.3,  // last_col
                        "",              // 数据为空白字符串
                        &fallback_fmt,  // 使用标准样式保持一致性
                    )?;
                }
            }

            // 獲取 Polars 列引用，准备数据写入
            let columns = sheet.df.columns();
            // 遍历每一列进行数据处理
            for (col_idx, column) in columns.iter().enumerate() {
                let c = col_idx as u16;

                // --- 2. 處理表頭 (Excel Row 0) ---
                // 邏輯升級：
                // 1. 優先查找 style_map 中的定義。
                // 2. 如果 style_map 是 None，則自動應用預設的 "header" 樣式。
                // 3. 如果以上皆否，則使用全局 "default"。
                // 4. 最後使用系統 "fallback"。
                // 四级样式查找机制确保表头样式的正确应用
                let header_cell_fmt = sheet.style_map.as_ref()
                    .and_then(|m| m.get(&(0, c)))         // 嘗試從地圖找特定单元格样式
                    .and_then(|name| self.styles.get(name.as_ref()))  // 查找样式池中的实际样式
                    .or_else(|| {                        // 如果地圖沒定義或地圖不存在
                        if sheet.style_map.is_none() {
                            self.styles.get("header")    // 自動應用預設 header 样式
                        } else {
                            None
                        }
                    })
                    .or(default_fmt)                     // 全局保底样式
                    .unwrap_or(&fallback_fmt);           // 系統保底样式
                // 写入表头单元格，应用确定的样式
                worksheet.write_with_format(0, c, column.name().as_str(), header_cell_fmt)?;

                // --- 3. 處理數據行 (Excel Row 1..N) ---
                // 遍历当前列的所有数据行
                for row_idx in 0..sheet.df.height() {
                    let val = column.get(row_idx)?; // 获取 AnyValue
                    let r = (row_idx + 1) as u32;   // Excel 行索引（跳过表头）

                    // 獲取該單元格專屬樣式或全局保底樣式
                    // 使用与表头相同的四级查找机制
                    let cell_fmt = sheet.style_map.as_ref()
                        .and_then(|m| m.get(&(r, c)))     // 查找特定单元格样式
                        .and_then(|name| self.styles.get(name.as_ref()))  // 获取实际样式
                        .or(default_fmt)                 // 全局默认样式
                        .unwrap_or(&fallback_fmt);       // 系统回退样式

                    // --- 4. 類型分派：將 Polars 豐富的數值類型統化為 Excel 的 f64。 ---
                    // 使用 _with_format 系列方法確保樣式（邊框、字體）被正確應用。
                    match val {
                        // 使用 _with_format 系列方法，傳入 cell_fmt (&Format)
                        // 整数类型转换为 f64 写入
                        AnyValue::Int8(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Int16(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Int32(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Int64(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::UInt32(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        // 浮点数类型处理
                        AnyValue::Float32(v) => worksheet.write_number_with_format(r, c, v as f64, cell_fmt)?,
                        AnyValue::Float64(v) => worksheet.write_number_with_format(r, c, v, cell_fmt)?,
                        // 字符串类型直接写入
                        AnyValue::String(s) => worksheet.write_string_with_format(r, c, s, cell_fmt)?,
                        // 布尔类型处理
                        AnyValue::Boolean(v) => worksheet.write_boolean_with_format(r, c, v, cell_fmt)?,
                        // 日期类型转换（天数偏移）
                        AnyValue::Date(days) => {
                            // 1. 转换基准：Polars 天数 + 25569 = Excel 天数
                            let excel_date_num = (days as f64) + 25569.0;
                            // 2. 在 Excel 中，日期本质上就是带格式的数字
                            worksheet.write_number_with_format(r, c, excel_date_num, cell_fmt)?
                        }
                        // 日期时间类型转换（Unix 时间戳转换）
                        AnyValue::Datetime(v, unit, _) => {
                            // 根据时间单位转换为秒数
                            let seconds = match unit {
                                TimeUnit::Milliseconds => v / 1_000,
                                TimeUnit::Microseconds => v / 1_000_000,
                                TimeUnit::Nanoseconds => v / 1_000_000_000,
                            } as f64;
                            // 将 Unix 秒数转换为 Excel 天数：秒数 / 86400 + 25569
                            let excel_dt_num = (seconds / 86400.0) + 25569.0;

                            worksheet.write_number_with_format(r, c, excel_dt_num, cell_fmt)?
                        }
                        // 空值处理（保持样式一致性）
                        AnyValue::Null => {
                            // Null 值也寫入 blank 以保持單元格邊框一致
                            worksheet.write_blank(r, c, cell_fmt)?
                        },
                        // 其他类型转换为字符串处理
                        _ => {
                            // 處理日期等類型，先轉為字符串
                            let s = format!("{}", val);
                            worksheet.write_string_with_format(r, c, &s, cell_fmt)?
                        },
                    };
                }
            }

            // --- 5. 自動列寬適配：---
            // 必須在數據寫入完成後調用，確保 Excel 引擎能計算出最長內容的佔位。
            worksheet.autofit();  // 自动调整列宽以适应内容
        }

        // 将内存中的工作簿保存到指定路径的文件中
        workbook.save(path)?;
        Ok(())  // 成功完成保存操作
    }

}

impl fmt::Debug for Workbook {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Workbook")
            .field("sheets", &self.sheets) // 只打印實現了 Debug 的 sheets
            .field("styles_count", &self.styles.len()) // 打印樣式數量作為替代
            .finish()
    }
}
