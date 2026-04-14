//! Region 工厂模块
//!
//! 本模块提供统一的工厂接口，用于整合多个工厂（StyleFactory、MergeFactory、DimensionFactory 等）
//! 生成完整的 RegionStyles。

use crate::dimension_factory::DimensionFactory;
use crate::merge_factory::MergeFactory;
use crate::region_styles::RegionStyles;
use crate::style_factory::StyleFactory;
use polars::prelude::DataFrame;
use std::error::Error;

/// Region 工厂
///
/// 整合多个工厂，生成完整的 RegionStyles。
/// 这是一个聚合层，本身不实现具体的样式或合并逻辑，
/// 而是协调各个子工厂的输出。
pub struct RegionFactory {
    /// 样式工厂（可选）
    style_factory: Option<StyleFactory>,
    /// 合并工厂（可选）
    merge_factory: Option<MergeFactory>,
    /// 维度工厂（可选）
    dimension_factory: Option<DimensionFactory>,
}

impl RegionFactory {
    /// 创建空的 RegionFactory
    ///
    /// 不包含任何子工厂，执行时返回空的 RegionStyles。
    pub fn new() -> Self {
        Self {
            style_factory: None,
            merge_factory: None,
            dimension_factory: None,
        }
    }

    /// 从 JSON 配置创建 RegionFactory
    ///
    /// # 配置格式
    /// ```json
    /// {
    ///   "style_rules": [...],
    ///   "merge_rules": [...],
    ///   "dimension_rules": [...]
    /// }
    /// ```
    ///
    /// # 参数
    /// * `config` - JSON 配置值
    ///
    /// # 返回值
    /// * `Ok(RegionFactory)` - 成功创建
    /// * `Err(Box<dyn Error>)` - 配置解析错误
    pub fn from_json(config: &serde_json::Value) -> Result<Self, Box<dyn Error>> {
        let style_factory = config
            .get("style_rules")
            .map(|r| StyleFactory::new(r.clone()))
            .transpose()?;

        let merge_factory = config
            .get("merge_rules")
            .map(|r| MergeFactory::new(r.clone()))
            .transpose()?;

        let dimension_factory = config
            .get("dimension_rules")
            .map(|r| DimensionFactory::new(r.clone()))
            .transpose()?;

        Ok(Self {
            style_factory,
            merge_factory,
            dimension_factory,
        })
    }

    /// 设置样式工厂
    pub fn with_style_factory(mut self, factory: StyleFactory) -> Self {
        self.style_factory = Some(factory);
        self
    }

    /// 设置合并工厂
    pub fn with_merge_factory(mut self, factory: MergeFactory) -> Self {
        self.merge_factory = Some(factory);
        self
    }

    /// 设置维度工厂
    pub fn with_dimension_factory(mut self, factory: DimensionFactory) -> Self {
        self.dimension_factory = Some(factory);
        self
    }

    /// 执行所有工厂，生成完整的 RegionStyles
    ///
    /// # 执行流程
    /// 1. 如果有 style_factory，执行它填充 cell_styles
    /// 2. 如果有 merge_factory，执行它填充 merge_ranges
    /// 3. 如果有 dimension_factory，执行它填充 row_heights 和 col_widths
    /// 4. 返回完整的 RegionStyles
    ///
    /// # 参数
    /// * `df` - 输入的 DataFrame
    ///
    /// # 返回值
    /// * `Ok(RegionStyles)` - 完整的样式属性
    /// * `Err(Box<dyn Error>)` - 执行过程中的错误
    pub fn execute(&self, df: &DataFrame) -> Result<RegionStyles, Box<dyn Error>> {
        let mut styles = RegionStyles::new();

        // 1. 执行样式工厂（填充 cell_styles）
        if let Some(factory) = &self.style_factory {
            styles.cell_styles = factory.execute(df)?;
        }

        // 2. 执行合并工厂（填充 merge_ranges）
        if let Some(factory) = &self.merge_factory {
            styles.merge_ranges = factory.execute(df)?;
        }

        // 3. 执行维度工厂（填充 row_heights 和 col_widths）
        if let Some(factory) = &self.dimension_factory {
            let dimension_result = factory.execute(df)?;
            styles.row_heights = dimension_result.row_heights;
            styles.col_widths = dimension_result.col_widths;
        }

        Ok(styles)
    }
}

impl Default for RegionFactory {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn test_region_factory_new() {
        let factory = RegionFactory::new();
        assert!(factory.style_factory.is_none());
        assert!(factory.merge_factory.is_none());
        assert!(factory.dimension_factory.is_none());
    }

    #[test]
    fn test_region_factory_from_json() {
        let config = json!({
            "style_rules": [
                {
                    "row_conditions": [{"type": "index", "criteria": [0]}],
                    "apply": {"style": "header"}
                }
            ],
            "merge_rules": [
                {"type": "vertical_match", "targets": ["部门"]}
            ],
            "dimension_rules": [
                {
                    "target": "row",
                    "condition": {"type": "index", "criteria": [0]},
                    "value": {"type": "fixed", "value": 30.0}
                }
            ]
        });

        let factory = RegionFactory::from_json(&config).unwrap();
        assert!(factory.style_factory.is_some());
        assert!(factory.merge_factory.is_some());
        assert!(factory.dimension_factory.is_some());
    }

    #[test]
    fn test_region_factory_with_dimension_factory() {
        let dimension_factory = DimensionFactory::from_json_str(
            r#"[{"target": "row", "condition": {"type": "index", "criteria": [0]}, "value": {"type": "fixed", "value": 30.0}}]"#
        ).unwrap();

        let factory = RegionFactory::new()
            .with_dimension_factory(dimension_factory);
        
        assert!(factory.dimension_factory.is_some());
    }
}
