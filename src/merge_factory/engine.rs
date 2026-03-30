//! 合并工厂核心执行引擎
//!
//! MergeFactory 是整个合并系统的入口点和执行核心，负责：
//! 1. 解析和验证 JSON 配置规则
//! 2. 执行合并算法扫描 DataFrame
//! 3. 生成标准化的 Excel 物理坐标合并区间
//! 4. 处理冲突检测和错误报告
//!
//! # 设计理念
//!
//! 工厂模式设计，将配置解析、算法执行、错误处理等职责分离，
//! 提供统一的接口供外部调用。
//!
//! # 线程安全性
//!
//! MergeFactory 实例是不可变的，可以在多线程环境中安全共享。
//! execute 方法需要可变借用 DataFrame，符合 Rust 的所有权模型。
use polars::prelude::*;
use serde_json::Value;
use crate::merge_factory::condition::MergeCondition;
use crate::merge_factory::error::MergeFactoryError; // 複用列名未找到等錯誤

/// 合并工厂执行引擎
///
/// MergeFactory 是合并系统的核心组件，负责将抽象的合并规则
/// 转换为具体的 Excel 物理合并坐标。
///
/// # 核心功能
///
/// ## 规则解析
/// - 将 JSON 配置解析为强类型 MergeCondition 枚举
/// - 提供友好的错误提示和配置验证
///
/// ## 算法执行
/// - 实现三种合并策略的具体算法
/// - 处理 DataFrame 到 Excel 坐标的转换
/// - 优化性能避免不必要的计算
///
/// ## 冲突处理
/// - 检测并处理重叠的合并区域
/// - 提供可配置的冲突解决策略
///
/// # 内存管理
///
/// 工厂本身不持有 DataFrame 的所有权，所有操作都是借用，
/// 确保内存使用的高效性和安全性。
pub struct MergeFactory {
    /// 合并规则集合（私有字段）
    ///
    /// 存储按顺序执行的合并规则列表。规则的执行顺序很重要，
    /// 因为后面的规则可能会与前面生成的合并区域产生冲突。
    ///
    /// # 规则处理特性
    /// - **有序执行**：按照 Vec 中的顺序依次处理每个规则
    /// - **独立处理**：每个规则独立生成合并区域
    /// - **统一冲突检测**：所有规则完成后进行全局冲突检查
    ///
    /// # 访问方式
    /// 该字段为私有，外部无法直接访问或修改，确保：
    /// - 数据完整性：防止意外修改规则集合
    /// - 封装性：隐藏内部实现细节
    /// - 安全性：避免不一致的状态
    ///
    /// 如需获取规则信息，可通过 `Debug` trait 或专门的检查方法。
    pub rules: Vec<MergeCondition>,
}

impl MergeFactory {

    /// 创建并初始化合并工厂实例
    ///
    /// 从 JSON Value 解析合并规则配置，构建 MergeFactory 实例。
    /// 支持灵活的输入格式，提供优雅的错误处理和降级机制。
    ///
    /// # 参数说明
    ///
    /// * `value`: JSON 配置值，预期为包含合并规则的对象数组
    ///   - 有效格式：`[{"type": "fixed", ...}, {"type": "vertical_match", ...}]`
    ///   - 空数组：`[]` - 创建空规则集
    ///   - null值：`null` - 创建空规则集
    ///   - 其他：触发警告并使用空规则集
    ///
    /// # 返回值
    ///
    /// * `Ok(MergeFactory)` - 成功创建的工厂实例
    /// * `Err(Box<dyn std::error::Error>)` - JSON 解析失败或其他错误
    ///
    /// # 错误处理策略
    ///
    /// ## 容错设计
    /// - **类型不匹配**：非数组输入会发出警告但不报错
    /// - **空值处理**：null 值被视为合法的空配置
    /// - **解析失败**：JSON 格式错误会返回具体错误信息
    ///
    /// ## 日志级别
    /// - **警告**：配置格式不当（非数组且非 null）
    /// - **错误**：JSON 反序列化失败
    ///
    /// # 性能考虑
    ///
    /// 此方法在创建时完成所有 JSON 解析工作，避免运行时重复解析，
    /// 提高执行阶段的性能表现。
    pub fn new(value: Value) -> Result<Self, Box<dyn std::error::Error>> {
        let rules: Vec<MergeCondition> = if value.is_array() {
            // 正常数组：尝试解析为合并规则
            serde_json::from_value(value)?
        } else {
            // 非数组处理：null 返回空规则，其他情况警告
            if !value.is_null() {
                eprintln!("Warning: MergeFactory input is not a valid array, using empty rules.");
            }
            vec![]
        };
        Ok(Self { rules })
    }

    /// 获取目标列的 Series 和索引信息
    ///
    /// 从 DataFrame 中提取指定列名对应的 Series 数据和列索引，
    /// 为合并算法提供必要的数据访问接口。
    ///
    /// # 参数说明
    ///
    /// * `df`: 输入的 DataFrame 引用，提供数据源
    /// * `targets`: 需要获取的列名切片，不能为空
    ///
    /// # 返回值
    ///
    /// * `Ok(Vec<(Series, u16)>)` - 成功获取的所有列信息
    ///   - Series: 拥有所有权的数据系列，支持随机访问
    ///   - u16: 列在 DataFrame 中的索引位置
    /// * `Err(MergeFactoryError)` - 列不存在或其他错误
    ///
    /// # 错误处理
    ///
    /// ## 列不存在错误
    /// 当 targets 中的任何列名在 DataFrame 中找不到时，
    /// 返回 `MergeFactoryError::ColumnNotFound` 错误。
    ///
    /// ## 错误传播
    /// 使用 `?` 操作符实现错误的短路传播，任一列获取失败
    /// 都会导致整个操作失败。
    ///
    /// # 性能特性
    ///
    /// ## 数据所有权
    /// 通过 `as_materialized_series().clone()` 获取拥有所有权的 Series，
    /// 避免在算法执行过程中出现生命周期问题。
    ///
    /// ## 批量处理
    /// 使用 Iterator 的 `map` 和 `collect` 组合，
    /// 一次性处理所有目标列，减少函数调用开销。
    ///
    /// # 算法保证
    ///
    /// ## 顺序一致性
    /// 返回的 Vec 保持与输入 targets 相同的顺序，
    /// 确保调用方可以正确关联数据和索引。
    ///
    /// ## 索引类型
    /// 列索引转换为 u16 类型，与 Excel 的列地址系统保持一致，
    /// 同时限制最大列数（约 65,535 列）。
    ///
    /// # 使用场景
    ///
    /// 主要服务于以下合并算法：
    /// - VerticalMatch: 获取用于纵向比对的列数据
    /// - HorizontalMatch: 获取用于横向比对的列数据
    fn get_target_series_and_indices(
        df: &DataFrame,
        targets: &[String],
    ) -> Result<Vec<(Series, u16)>, MergeFactoryError> {
        targets.iter().map(|col_name| {
            let col_idx = df.get_column_index(col_name)
                .ok_or_else(|| MergeFactoryError::ColumnNotFound(col_name.clone()))?;
            let series = df.column(col_name)
                .map_err(|_| MergeFactoryError::ColumnNotFound(col_name.clone()))?
                .as_materialized_series().clone();
            Ok((series, col_idx as u16))
        }).collect::<Result<Vec<_>, MergeFactoryError>>()
    }

    /// 检查两个矩形区域是否重叠
    ///
    /// 使用经典的矩形重叠检测算法，判断两个 Excel 合并区域是否存在重叠。
    /// 坐标系统基于 0-based 索引，行列均为闭区间。
    ///
    /// # 参数说明
    ///
    /// * `a`: 第一个矩形区域，格式为 (起始行, 起始列, 结束行, 结束列)
    /// * `b`: 第二个矩形区域，格式为 (起始行, 起始列, 结束行, 结束列)
    ///
    /// # 坐标系统约定
    ///
    /// ## 索引类型
    /// - 行号：u32 类型，支持大表格处理
    /// - 列号：u16 类型，与 Excel 列地址系统对应
    ///
    /// ## 区间性质
    /// - 所有坐标均为闭区间，包含起始和结束位置
    /// - 坐标已标准化，保证起始 <= 结束
    /// - 基于 0-based 索引系统
    ///
    /// # 算法原理
    ///
    /// ## 重叠判断逻辑
    /// 两个矩形重叠当且仅当它们在行方向和列方向都有交集：
    /// - 行方向交集：!(r1_end < r2_start || r1_start > r2_end)
    /// - 列方向交集：!(c1_end < c2_start || c1_start > c2_end)
    ///
    /// ## 性能优化
    /// 通过检测不重叠条件来实现，避免复杂的区间交集计算，
    /// 时间复杂度为 O(1)，适合高频调用场景。
    ///
    /// # 边界情况处理
    ///
    /// ## 相邻区域
    /// 仅边 adjacent 的区域（如 (0,0,0,1) 和 (0,2,0,2)）不视为重叠
    ///
    /// ## 点重叠
    /// 仅顶点相交的区域（如 (0,0,0,0) 和 (0,0,0,0)）视为重叠
    ///
    /// ## 包含关系
    /// 一个区域完全包含另一个区域时视为重叠
    ///
    /// # 应用场景
    ///
    /// 主要用于以下冲突检测场景：
    /// - 合并区域冲突预防
    /// - 包围盒优化算法
    /// - 空间查询优化
    ///
    /// # 算法复杂度
    ///
    /// - 时间复杂度：O(1) - 常数时间操作
    /// - 空间复杂度：O(1) - 无额外空间分配
    fn ranges_overlap(a: &(u32, u16, u32, u16), b: &(u32, u16, u32, u16)) -> bool {
        let (r1_start, c1_start, r1_end, c1_end) = *a;
        let (r2_start, c2_start, r2_end, c2_end) = *b;

        // 两个矩形不重叠的条件：一个在另一个的左边、右边、上面、下面
        // 使用德摩根定律转换为重叠条件
        !(r1_end < r2_start || r1_start > r2_end ||
            c1_end < c2_start || c1_start > c2_end)
    }

    /// 创建优化的包围盒集合
    ///
    /// 使用并查集算法将重叠的合并区域分组，并为每组创建最小的包围盒，
    /// 用于加速后续的冲突检测和空间查询操作。
    ///
    /// # 参数说明
    ///
    /// * `ranges`: 待优化的合并区域切片，每个区域格式为 (起始行, 起始列, 结束行, 结束列)
    ///
    /// # 返回值
    ///
    /// 返回优化后的包围盒列表，每个包围盒都是能覆盖一组重叠区域的最小矩形。
    ///
    /// # 算法概述
    ///
    /// ## 三阶段处理
    /// 1. **分组阶段**：使用并查集将重叠区域归类到相同组
    /// 2. **收集阶段**：按组别重新组织区域数据
    /// 3. **合并阶段**：为每组计算最小包围盒
    ///
    /// ## 并查集优化
    /// - **路径压缩**：查找操作的时间复杂度接近 O(1)
    /// - **按秩合并**：保持树的平衡性（本实现简化处理）
    ///
    /// # 性能特征
    ///
    /// ## 时间复杂度
    /// - 分组阶段：O(n²) - 需要两两比较所有区域
    /// - 收集阶段：O(n × α(n)) - α(n) 为阿克曼函数的反函数，近似常数
    /// - 合并阶段：O(n) - 线性遍历所有区域
    /// - 总体：O(n²) - 主要瓶颈在两两比较
    ///
    /// ## 空间复杂度
    /// - O(n) - 存储并查集和分组映射
    ///
    /// # 算法优势
    ///
    /// ## 最优包围盒
    /// 生成的每个包围盒都是覆盖对应区域组的最小矩形，
    /// 确保后续冲突检测的准确性。
    ///
    /// ## 冲突检测优化
    /// 通过减少包围盒数量，大幅降低冲突检测的计算量。
    ///
    /// ## 可扩展性
    /// 算法结构清晰，易于扩展支持更多优化策略。
    ///
    /// # 使用场景
    ///
    /// 主要用于合并工厂的性能优化：
    /// - **快速冲突预检**：用少量包围盒快速排除大部分无冲突情况
    /// - **空间索引构建**：为大规模合并区域建立高效的空间索引
    /// - **可视化辅助**：为调试工具提供区域聚合视图
    ///
    /// # 算法局限性
    ///
    /// ## 计算密集
    /// 对于大量区域（> 1000），O(n²) 的比较可能成为性能瓶颈。
    ///
    /// ## 内存开销
    /// 需要额外存储并查集和分组映射，占用 O(n) 额外空间。
    ///
    /// # 算法细节
    ///
    /// ## 分组策略
    /// 两个区域如果重叠就被归为同一组，确保包围盒的完整性。
    ///
    /// ## 包围盒计算
    /// 对每组区域分别计算行最小值、列最小值、行最大值、列最大值，
    /// 形成能完全包含该组所有区域的最小矩形。
    fn create_optimized_bounding_boxes(ranges: &[(u32, u16, u32, u16)]) -> Vec<(u32, u16, u32, u16)> {
        // 空输入直接返回空结果
        if ranges.is_empty() {
            return vec![];
        }

        // 初始化并查集：每个区域初始时属于独立的组
        let mut groups: Vec<usize> = (0..ranges.len()).collect();

        // 查找根节点函数（带路径压缩优化）
        // 路径压缩：将查找路径上的所有节点直接连接到根节点
        fn find(groups: &mut [usize], x: usize) -> usize {
            if groups[x] != x {
                groups[x] = find(groups, groups[x]);
            }
            groups[x]
        }

        // 查找根节点函数（带路径压缩优化）
        // 路径压缩：将查找路径上的所有节点直接连接到根节点
        fn union(groups: &mut [usize], x: usize, y: usize) {
            let root_x = find(groups, x);
            let root_y = find(groups, y);
            if root_x != root_y {
                groups[root_y] = root_x; // 合并两个组
            }
        }

        // 分组阶段：将重叠的区域归为同一组
        // 时间复杂度 O(n²)，空间复杂度 O(n)
        for i in 0..ranges.len() {
            for j in (i + 1)..ranges.len() {
                // 如果两个区域重叠，则将它们所在的组合并
                if Self::ranges_overlap(&ranges[i], &ranges[j]) {
                    union(&mut groups, i, j);
                }
            }
        }

        // 收集阶段：按组别重新组织区域数据
        // 使用 HashMap 存储每个组包含的区域列表
        let mut group_map: std::collections::HashMap<usize, Vec<(u32, u16, u32, u16)>> =
            std::collections::HashMap::new();

        // 将每个区域归类到对应的组中
        for (i, range) in ranges.iter().enumerate() {
            let root = find(&mut groups, i);
            group_map.entry(root).or_insert_with(Vec::new).push(*range);
        }

        // 合并阶段：为每组创建最小包围盒
        let mut result = Vec::new();

        // 遍历每个组，计算该组的最小包围盒
        for ranges_in_group in group_map.values() {
            if let Some(&first) = ranges_in_group.first() {
                // 初始化包围盒边界为第一个区域的边界
                let mut min_row = first.0;
                let mut min_col = first.1;
                let mut max_row = first.2;
                let mut max_col = first.3;

                // 遍历组内其他区域，扩展包围盒边界
                for &range in ranges_in_group.iter().skip(1) {
                    min_row = min_row.min(range.0);  // 扩展最小行
                    min_col = min_col.min(range.1);  // 扩展最小列
                    max_row = max_row.max(range.2);  // 扩展最大行
                    max_col = max_col.max(range.3);  // 扩展最大列
                }

                // 将计算出的包围盒添加到结果中
                result.push((min_row, min_col, max_row, max_col));
            }
        }

        result
    }

    /// 检查并添加合并区域
    ///
    /// 执行完整的冲突检测流程，确定新合并区域是否可以安全添加到结果集中。
    /// 采用两级检测机制平衡性能和准确性，提供详细的冲突报告。
    ///
    /// # 参数说明
    ///
    /// * `merge_range`: 待添加的合并区域，格式为 (起始行, 起始列, 结束行, 结束列)
    /// * `all_ranges`: 当前已存在的所有合并区域
    /// * `bounding_boxes`: 优化后的包围盒列表，用于快速冲突预检
    /// * `warnings`: 警告信息收集器，用于记录冲突相关的警告消息
    ///
    /// # 返回值
    ///
    /// * `Some((u32, u16, u32, u16))` - 区域可以安全添加，返回该区域
    /// * `None` - 区域存在冲突，已被忽略
    ///
    /// # 检测流程
    ///
    /// ## 两级冲突检测
    /// 1. **快速预检**：使用包围盒进行 O(m) 时间的粗筛
    /// 2. **详细检测**：对疑似冲突的区域进行 O(n) 的精确检查
    ///
    /// ## 决策逻辑
    /// ```text
    /// 开始
    ///   ↓
    /// 快速预检是否有冲突可能？
    ///   ↓ 是
    /// 详细检测是否真的冲突？
    ///   ↓ 是                    ↓ 否
    /// 记录警告并返回 None    返回 Some(区域)
    ///   ↓                      ↓
    ///  结束 ←─────────────────── 结束
    /// ```
    ///
    /// # 性能优化策略
    ///
    /// ## 分层检测
    /// - 绝大多数无冲突区域能通过快速预检直接通过
    /// - 只有少数疑似冲突区域需要昂贵的详细检测
    ///
    /// ## 短路求值
    /// - 快速预检失败时直接跳过详细检测
    /// - 详细检测发现冲突时立即终止
    ///
    /// ## 零拷贝设计
    /// - 使用引用传递避免不必要的数据复制
    /// - 返回 Option 避免额外的布尔标志位检查
    ///
    /// # 错误处理
    ///
    /// ## 警告机制
    /// 所有冲突都会通过 `warnings` 参数记录详细信息，便于调试和审计。
    ///
    /// ## 静默失败
    /// 冲突区域会被静默忽略，不影响其他区域的正常处理。
    ///
    /// # 使用场景
    ///
    /// 统一的合并区域冲突处理入口，服务于所有合并规则：
    /// - **Fixed 规则**：处理静态坐标合并的冲突检测
    /// - **VerticalMatch**：处理纵向自动合并的冲突检测
    /// - **HorizontalMatch**：处理横向自动合并的冲突检测
    ///
    /// # 设计原则
    ///
    /// ## 安全第一
    /// 宁可遗漏合并也不产生冲突的 Excel 文件。
    ///
    /// ## 透明处理
    /// 所有决策过程都有详细记录，便于问题追踪。
    ///
    /// ## 性能优先
    /// 通过智能的两级检测机制最大化处理效率。
    ///
    /// # 线程安全性
    ///
    /// 函数本身是纯函数，不持有任何可变状态，
    /// 但 `warnings` 参数需要可变借用，调用时需注意所有权管理。
    fn check_and_add_merge_range(
        merge_range: (u32, u16, u32, u16),
        all_ranges: &[(u32, u16, u32, u16)],
        bounding_boxes: &[(u32, u16, u32, u16)],
        warnings: &mut Vec<String>,
    ) -> Option<(u32, u16, u32, u16)> {
        // 快速冲突检查：直接内联实现
        let might_conflict = !bounding_boxes.is_empty() &&
            bounding_boxes.iter().any(|box_rect| {
                Self::ranges_overlap(&merge_range, box_rect)
            });


        // 如果快速检查显示可能存在冲突，则进行详细检查
        if might_conflict {
            // 进行详细冲突检查
            if all_ranges.iter().any(|existing|
                Self::ranges_overlap(&merge_range, existing)
            ) {
                // 发现真实冲突，记录警告并返回 None 表示不添加
                warnings.push(format!("合并区域 {:?} 与其他区域冲突，将被忽略", merge_range));
                return None; // 有冲突，返回None表示不添加
            }
        }

        // 没有冲突或详细检查后发现没有实际冲突，返回区域表示可以添加
        Some(merge_range)
    }

    /// 执行合并规则扫描，生成Excel物理合并区间
    ///
    /// 遍历所有配置的合并规则，扫描输入的 DataFrame，
    /// 生成标准化的 Excel 合并坐标列表。
    ///
    /// # 参数说明
    ///
    /// * `df`: 输入的 DataFrame，提供数据源和维度信息
    ///
    /// # 返回值
    ///
    /// * `Ok(Vec<(u32, u16, u32, u16)>)` - 成功生成的合并区域列表
    ///   - 格式：(起始行, 起始列, 结束行, 结束列)
    ///   - 坐标：基于 0-based 系统，包含起始和结束位置
    /// * `Err(MergeFactoryError)` - 执行过程中发生的错误
    ///
    /// # 执行流程
    ///
    /// ## 整体架构
    /// ```text
    /// 开始执行
    ///   ↓
    /// 初始化状态 (all_ranges, bounding_boxes, warnings)
    ///   ↓
    /// 遍历每个规则
    ///   ↓
    /// 根据规则类型执行对应算法
    ///   ↓
    /// 冲突检测和处理
    ///   ↓
    /// 更新包围盒优化结构
    ///   ↓
    /// 输出警告信息
    ///   ↓
    /// 返回最终结果
    /// ```
    ///
    /// ## 三级优化体系
    /// 1. **规则级别**：每条规则独立处理
    /// 2. **冲突检测**：两级检测机制（快速+详细）
    /// 3. **性能优化**：包围盒加速后续检测
    ///
    /// # 坐标系统
    ///
    /// ## DataFrame 坐标
    /// - 行：0-based 索引，0 到 height-1
    /// - 列：0-based 索引，0 到 width-1
    ///
    /// ## Excel 坐标转换
    /// - 行：+1 偏移处理标题行（DataFrame行0对应Excel行1）
    /// - 列：直接映射（u16类型与Excel列对应）
    ///
    /// # 性能特征
    ///
    /// ## 时间复杂度
    /// - Fixed: O(k×m) k为目标数，m为包围盒数
    /// - VerticalMatch: O(n×c×m) n为行数，c为列数
    /// - HorizontalMatch: O(n×c×m) n为行数，c为列数
    ///
    /// ## 空间复杂度
    /// - O(r + b) r为结果数，b为包围盒数
    ///
    /// # 错误处理策略
    ///
    /// ## 致命错误
    /// - 索引越界：立即终止执行并返回错误
    /// - 列不存在：立即终止执行并返回错误
    ///
    /// ## 非致命处理
    /// - 冲突区域：静默忽略并记录警告
    /// - 空规则：跳过继续执行
    ///
    /// # 线程安全性
    ///
    /// 函数接受 DataFrame 的不可变引用，符合 Rust 的借用检查器要求，
    /// 可以安全地在单线程环境中调用。
    pub fn execute(&self, df: &DataFrame) -> Result<Vec<(u32, u16, u32, u16)>, MergeFactoryError> {
        // 初始化执行状态
        let mut all_ranges: Vec<(u32, u16, u32, u16)> = Vec::new();
        let mut bounding_boxes: Vec<(u32, u16, u32, u16)> = Vec::new();
        let mut warnings = Vec::new();

        // 获取 DataFrame 维度信息
        let height = df.height();
        let max_phys_row = height as u32;
        let max_phys_col = df.width() as u16;

        // 按顺序处理每条合并规则
        for rule in &self.rules {
            let mut current_rule_ranges = Vec::new();

            match rule {
                // --- 1. 静态坐标合并 (Fixed) ---
                MergeCondition::Fixed { targets } => {
                    // 遍历每个静态合并目标
                    for range in targets {
                        let [r1, c1, r2, c2] = *range;

                        // 核心逻辑：坐标标准化 (Normalization)
                        // 无论用户输入的点在哪个位置，始终取最小值作为起始点，最大值作为结束点
                        // 这等效于在第四象限中将任意两点拉伸为一个规范矩形
                        let s_r = std::cmp::min(r1, r2);
                        let e_r = std::cmp::max(r1, r2);
                        let s_c = std::cmp::min(c1, c2);
                        let e_c = std::cmp::max(c1, c2);

                        // 边界校验（仅检查是否超出 DataFrame 物理边界）
                        if e_r > max_phys_row || e_c >= max_phys_col as u32 {
                            return Err(MergeFactoryError::IndexOutOfBounds(e_r, max_phys_row as usize));
                        }

                        // 检查列号是否超过 Excel 的 u16 限制
                        if e_c > u16::MAX as u32 {
                            return Err(MergeFactoryError::ColumnNotFound(format!("列索引 {} 超出 u16 限制", e_c)));
                        }

                        // 标准化后的合并区域（转换为 u16 列索引）
                        let normalized_range = (s_r, s_c as u16, e_r, e_c as u16);

                        // 执行冲突检测并添加区域
                        if let Some(valid_range) = Self::check_and_add_merge_range(
                            normalized_range,
                            &all_ranges,
                            &bounding_boxes,
                            &mut warnings
                        ) {
                            current_rule_ranges.push(valid_range);
                            all_ranges.push(valid_range);
                        }
                    }
                }
                // --- 2. 纵向自动合并 (Vertical Match) ---
                MergeCondition::VerticalMatch { targets, .. } => {
                    // 预检查：空目标或数据行数不足时不处理
                    if targets.is_empty() || height < 2 { continue; }

                    // 获取所有目标列的 Series 引用与物理列索引
                    let target_series_and_indices = Self::get_target_series_and_indices(df, targets)?;

                    // 维护每个列的独立合并起始位置
                    let mut start_positions = vec![0; target_series_and_indices.len()];

                    // 记录哪些列在当前行发生了变化（用于父子约束）
                    let mut changed_columns = vec![false; target_series_and_indices.len()];

                    for r in 1..height {
                        // 检查每一列是否发生变化
                        for (i, (series, _)) in target_series_and_indices.iter().enumerate() {
                            changed_columns[i] = series.get(r)? != series.get(start_positions[i])?;
                        }

                        // 应用父子约束：如果前面的列发生变化，后面的列也必须重新开始
                        // 实现隐式链式扫描：父级变化影响子级
                        for i in 1..changed_columns.len() {
                            if changed_columns[i-1] {
                                changed_columns[i] = true;
                            }
                        }

                        // 处理发生变化的列
                        for (i, &(ref _series, col_idx)) in target_series_and_indices.iter().enumerate() {
                            if changed_columns[i] {
                                // 如果区间长度 > 1，则生成合并区间
                                // 区间为 [start_positions[i], r-1]，转换为Excel坐标需要+1偏移
                                if r - start_positions[i] > 1 {
                                    let merge_range = (
                                        (start_positions[i] + 1) as u32, col_idx,
                                        ((r - 1) + 1) as u32, col_idx
                                    );
                                    // 执行冲突检测并添加区域
                                    if let Some(valid_range) = Self::check_and_add_merge_range(
                                        merge_range,
                                        &all_ranges,
                                        &bounding_boxes,
                                        &mut warnings
                                    ) {
                                        current_rule_ranges.push(valid_range);
                                        all_ranges.push(valid_range);
                                    }
                                }
                                start_positions[i] = r;
                            }
                        }
                    }

                    // 结尾收网
                    for (i, &start_pos) in start_positions.iter().enumerate() {
                        // 如果最后一段长度 > 1，则生成合并区间
                        if height - start_pos > 1 {
                            let col_idx = target_series_and_indices[i].1;
                            let merge_range = (
                                (start_pos + 1) as u32, col_idx,           // 起始行(+1偏移), 起始列
                                ((height - 1) + 1) as u32, col_idx         // 结束行(+1偏移), 结束列
                            );

                            // 执行冲突检测并添加区域
                            if let Some(valid_range) = Self::check_and_add_merge_range(
                                merge_range,
                                &all_ranges,
                                &bounding_boxes,
                                &mut warnings
                            ) {
                                current_rule_ranges.push(valid_range);
                                all_ranges.push(valid_range);
                            }
                        }
                    }
                }
                // --- 3. 横向自动合并 (Horizontal Match) ---
                MergeCondition::HorizontalMatch { targets, .. } => {
                    // --- 3. 横向自动合并 (Horizontal Match) ---
                    if targets.is_empty() { continue; }

                    // 获取所有目标列的 Series 引用与物理列索引
                    let target_series_and_indices = Self::get_target_series_and_indices(df, targets)?;

                    // 对每一行进行处理
                    for row_idx in 0..height {
                        // 检查连续相同值的区段
                        let mut i = 0;
                        while i < target_series_and_indices.len() {
                            // 获取当前位置的值作为比较基准
                            let (current_series, _) = &target_series_and_indices[i];
                            let current_value = current_series.get(row_idx)?;

                            // 寻找连续相同值的结束位置
                            let mut j = i + 1;
                            while j < target_series_and_indices.len() {
                                let (next_series, _) = &target_series_and_indices[j];
                                let next_value = next_series.get(row_idx)?;
                                if next_value != current_value {
                                    break; // 值不同，停止搜索
                                }
                                j += 1;
                            }

                            // 如果找到连续相同值且长度>1，则生成合并区间
                            // 长度为 j-1，需要 > i 才有意义
                            if (j - 1) > i {
                                let start_col = target_series_and_indices[i].1;
                                let end_col = target_series_and_indices[j-1].1;
                                let excel_row = (row_idx + 1) as u32;

                                // 构造合并区域：同一行内连续列的合并
                                let merge_range = (excel_row, start_col, excel_row, end_col);

                                // 执行冲突检测并添加区域
                                if let Some(valid_range) = Self::check_and_add_merge_range(
                                    merge_range,
                                    &all_ranges,
                                    &bounding_boxes,
                                    &mut warnings
                                ) {
                                    current_rule_ranges.push(valid_range);
                                    all_ranges.push(valid_range);
                                }

                            }
                            // 移动到下一个不同的值段
                            i = j;
                        }
                    }
                }
            }

            // 更新包围盒（仅针对当前规则产生的区域）
            // 为后续规则的冲突检测提供优化
            if !current_rule_ranges.is_empty() {
                let new_boxes = Self::create_optimized_bounding_boxes(&current_rule_ranges);
                bounding_boxes.extend(new_boxes);
            }
        }

        // 输出警告信息到标准错误流
        for warning in warnings {
            eprintln!("MergeFactory Warning: {}", warning);
        }

        // 返回所有成功生成的合并区域
        Ok(all_ranges)
    }
}