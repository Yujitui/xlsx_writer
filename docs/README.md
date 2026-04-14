# xlsx_writer 文档

本文档库提供 xlsx_writer 库的全面使用指南。

## 📚 文档结构

### 主要文档

| 文档 | 说明 | 适用场景 |
|------|------|----------|
| [configuration_guide.md](./configuration_guide.md) | **完整配置指南** | 系统学习所有配置选项 |
| [QUICK_REFERENCE.md](./QUICK_REFERENCE.md) | **快速参考手册** | 日常开发快速查询 |

### 示例配置

| 示例 | 说明 |
|------|------|
| [basic.json](./examples/basic.json) | 基础表格样式 |
| [conditional.json](./examples/conditional.json) | 条件格式（数据条） |
| [merge.json](./examples/merge.json) | 合并单元格 |
| [dimensions.json](./examples/dimensions.json) | 行高列宽设置 |
| [complete.json](./examples/complete.json) | 完整复杂配置 |
| [sales_report.json](./examples/sales_report.json) | 实战示例：销售报表（含完整样式库） |

## 🚀 快速开始

### 1. 了解基础（5分钟）

阅读 [configuration_guide.md 第1-2章](./configuration_guide.md#1-快速开始)

### 2. 查看示例（10分钟）

浏览 [examples/](./examples/) 目录中的配置文件

### 3. 动手实践（30分钟）

参考 [QUICK_REFERENCE.md](./QUICK_REFERENCE.md) 创建自己的配置

## 📖 使用建议

**新手**：
1. 先阅读 [configuration_guide.md 第1-3章](./configuration_guide.md#1-快速开始)
2. 查看 [basic.json](./examples/basic.json) 和 [conditional.json](./examples/conditional.json)
3. 遇到问题查阅 [第9章 故障排除](./configuration_guide.md#9-故障排除)

**日常开发**：
1. 保存 [QUICK_REFERENCE.md](./QUICK_REFERENCE.md) 为书签
2. 需要具体示例时查看 [examples/](./examples/) 目录
3. 复杂场景参考 [complete.json](./examples/complete.json)

**进阶用户**：
1. 完整阅读 [configuration_guide.md](./configuration_guide.md)
2. 掌握所有 [6种条件类型](./configuration_guide.md#6-条件类型完整参考)
3. 学习 [常见用例 Cookbook](./configuration_guide.md#8-常见用例-cookbook)

## 🔧 核心概念

### 配置四大件

```json
{
  "styles": {                 // 样式库定义
    "header": { "bg_color": "#4472C4", "bold": true },
    "data": { "font_size": 10 }
  },
  "style_rules": [...],      // 单元格样式规则（引用 styles 中的样式）
  "merge_rules": [...],      // 合并单元格
  "dimension_rules": [...]   // 行高列宽
}
```

**关系说明**：
1. `styles` 定义所有可复用的样式模板
2. `style_rules` 通过样式名称引用 `styles` 中定义的样式
3. 在代码中使用 `with_library_from_json(&config["styles"])` 加载样式库

### 条件类型

- **Index**: 通过行/列索引选择
- **ValueRange**: 数值范围比较
- **Match**: 集合成员匹配
- **Find**: 字符串子串查找
- **Equal**: 列间相等比较
- **ExcludeRows**: 排除特定行

## 📞 获取帮助

如果在使用过程中遇到问题：

1. 查看 [故障排除章节](./configuration_guide.md#9-故障排除)
2. 检查 [常见错误速查](./QUICK_REFERENCE.md#4-常见错误速查)
3. 参考示例配置
4. 查看 API 文档

## 📝 文档统计

- **完整指南**: 1704 行，涵盖所有功能细节（包括样式库配置）
- **快速参考**: 368 行，常用模式速查
- **示例配置**: 6 个，覆盖主要使用场景

## 🔄 版本信息

- 文档版本: 1.0.0
- 对应库版本: 0.2.0+
- 最后更新: 2024年

---

**提示**: 建议将本文档库与项目代码一起保存，方便随时查阅。
