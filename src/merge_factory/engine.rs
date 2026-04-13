//! 合并工厂核心执行引擎
//!
//! MergeFactory 是整个合并系统的入口点和执行核心，负责：
//! 1. 解析和验证 JSON 配置规则
//! 2. 执行合并算法扫描 DataFrame
//! 3. 生成标准化的 Excel 物理坐标合并区间
//! 4. 处理冲突检测和错误报告
use crate::merge_factory::condition::MergeCondition;
use crate::merge_factory::error::MergeFactoryError;
use polars::prelude::*;
use serde_json::Value; // 複用列名未找到等錯誤

/// 合并工厂执行引擎
///
/// MergeFactory 是合并系统的核心组件，负责将抽象的合并规则
/// 转换为具体的 Excel 物理合并坐标。
pub struct MergeFactory {
    /// 合并规则集合（私有字段）
    ///
    /// 存储按顺序执行的合并规则列表。规则的执行顺序很重要，
    /// 因为后面的规则可能会与前面生成的合并区域产生冲突。
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
    fn get_target_series_and_indices(
        df: &DataFrame,
        targets: &[String],
    ) -> Result<Vec<(Series, u16)>, MergeFactoryError> {
        targets
            .iter()
            .map(|col_name| {
                let col_idx = df
                    .get_column_index(col_name)
                    .ok_or_else(|| MergeFactoryError::ColumnNotFound(col_name.clone()))?;
                let series = df
                    .column(col_name)
                    .map_err(|_| MergeFactoryError::ColumnNotFound(col_name.clone()))?
                    .as_materialized_series()
                    .clone();
                Ok((series, col_idx as u16))
            })
            .collect::<Result<Vec<_>, MergeFactoryError>>()
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
    fn ranges_overlap(a: &(u32, u16, u32, u16), b: &(u32, u16, u32, u16)) -> bool {
        let (r1_start, c1_start, r1_end, c1_end) = *a;
        let (r2_start, c2_start, r2_end, c2_end) = *b;

        // 两个矩形不重叠的条件：一个在另一个的左边、右边、上面、下面
        // 使用德摩根定律转换为重叠条件
        !(r1_end < r2_start || r1_start > r2_end || c1_end < c2_start || c1_start > c2_end)
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
    fn create_optimized_bounding_boxes(
        ranges: &[(u32, u16, u32, u16)],
    ) -> Vec<(u32, u16, u32, u16)> {
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
                    min_row = min_row.min(range.0); // 扩展最小行
                    min_col = min_col.min(range.1); // 扩展最小列
                    max_row = max_row.max(range.2); // 扩展最大行
                    max_col = max_col.max(range.3); // 扩展最大列
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
    fn check_and_add_merge_range(
        merge_range: (u32, u16, u32, u16),
        all_ranges: &[(u32, u16, u32, u16)],
        bounding_boxes: &[(u32, u16, u32, u16)],
        warnings: &mut Vec<String>,
    ) -> Option<(u32, u16, u32, u16)> {
        // 快速冲突检查：直接内联实现
        let might_conflict = !bounding_boxes.is_empty()
            && bounding_boxes
                .iter()
                .any(|box_rect| Self::ranges_overlap(&merge_range, box_rect));

        // 如果快速检查显示可能存在冲突，则进行详细检查
        if might_conflict {
            // 进行详细冲突检查
            if all_ranges
                .iter()
                .any(|existing| Self::ranges_overlap(&merge_range, existing))
            {
                // 发现真实冲突，记录警告并返回 None 表示不添加
                warnings.push(format!(
                    "合并区域 {:?} 与其他区域冲突，将被忽略",
                    merge_range
                ));
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
                            return Err(MergeFactoryError::IndexOutOfBounds(
                                e_r,
                                max_phys_row as usize,
                            ));
                        }

                        // 检查列号是否超过 Excel 的 u16 限制
                        if e_c > u16::MAX as u32 {
                            return Err(MergeFactoryError::ColumnNotFound(format!(
                                "列索引 {} 超出 u16 限制",
                                e_c
                            )));
                        }

                        // 标准化后的合并区域（转换为 u16 列索引）
                        let normalized_range = (s_r, s_c as u16, e_r, e_c as u16);

                        // 执行冲突检测并添加区域
                        if let Some(valid_range) = Self::check_and_add_merge_range(
                            normalized_range,
                            &all_ranges,
                            &bounding_boxes,
                            &mut warnings,
                        ) {
                            current_rule_ranges.push(valid_range);
                            all_ranges.push(valid_range);
                        }
                    }
                }
                // --- 2. 纵向自动合并 (Vertical Match) ---
                MergeCondition::VerticalMatch { targets, .. } => {
                    // 预检查：空目标或数据行数不足时不处理
                    if targets.is_empty() || height < 2 {
                        continue;
                    }

                    // 获取所有目标列的 Series 引用与物理列索引
                    let target_series_and_indices =
                        Self::get_target_series_and_indices(df, targets)?;

                    // 维护每个列的独立合并起始位置
                    let mut start_positions = vec![0; target_series_and_indices.len()];

                    // 记录哪些列在当前行发生了变化（用于父子约束）
                    let mut changed_columns = vec![false; target_series_and_indices.len()];

                    for r in 1..height {
                        // 检查每一列是否发生变化
                        for (i, (series, _)) in target_series_and_indices.iter().enumerate() {
                            changed_columns[i] =
                                series.get(r)? != series.get(start_positions[i])?;
                        }

                        // 应用父子约束：如果前面的列发生变化，后面的列也必须重新开始
                        // 实现隐式链式扫描：父级变化影响子级
                        for i in 1..changed_columns.len() {
                            if changed_columns[i - 1] {
                                changed_columns[i] = true;
                            }
                        }

                        // 处理发生变化的列
                        for (i, &(ref _series, col_idx)) in
                            target_series_and_indices.iter().enumerate()
                        {
                            if changed_columns[i] {
                                // 如果区间长度 > 1，则生成合并区间
                                // 区间为 [start_positions[i], r-1]，转换为Excel坐标需要+1偏移
                                if r - start_positions[i] > 1 {
                                    let merge_range = (
                                        (start_positions[i] + 1) as u32,
                                        col_idx,
                                        ((r - 1) + 1) as u32,
                                        col_idx,
                                    );
                                    // 执行冲突检测并添加区域
                                    if let Some(valid_range) = Self::check_and_add_merge_range(
                                        merge_range,
                                        &all_ranges,
                                        &bounding_boxes,
                                        &mut warnings,
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
                                (start_pos + 1) as u32,
                                col_idx, // 起始行(+1偏移), 起始列
                                ((height - 1) + 1) as u32,
                                col_idx, // 结束行(+1偏移), 结束列
                            );

                            // 执行冲突检测并添加区域
                            if let Some(valid_range) = Self::check_and_add_merge_range(
                                merge_range,
                                &all_ranges,
                                &bounding_boxes,
                                &mut warnings,
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
                    if targets.is_empty() {
                        continue;
                    }

                    // 获取所有目标列的 Series 引用与物理列索引
                    let target_series_and_indices =
                        Self::get_target_series_and_indices(df, targets)?;

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
                                let end_col = target_series_and_indices[j - 1].1;
                                let excel_row = (row_idx + 1) as u32;

                                // 构造合并区域：同一行内连续列的合并
                                let merge_range = (excel_row, start_col, excel_row, end_col);

                                // 执行冲突检测并添加区域
                                if let Some(valid_range) = Self::check_and_add_merge_range(
                                    merge_range,
                                    &all_ranges,
                                    &bounding_boxes,
                                    &mut warnings,
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
