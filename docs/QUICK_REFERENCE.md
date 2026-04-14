# xlsx_writer 配置快速参考

本文档提供常用配置模式的快速查询。

## 目录

1. [条件类型速查](#1-条件类型速查)
2. [规则配置模板](#2-规则配置模板)
3. [样式属性速查](#3-样式属性速查)
4. [常见错误速查](#4-常见错误速查)

---

## 1. 条件类型速查

### Index（索引）

```json
// 单行
{"type": "index", "criteria": [0]}

// 多行
{"type": "index", "criteria": [0, 1, 2]}

// 最后一行
{"type": "index", "criteria": [-1]}

// 范围（配合排除使用）
{"type": "index", "criteria": [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]}
```

### ValueRange（数值范围）

```json
// 大于
{"type": "value_range", "targets": ["销售额"], "criteria": ">10000"}

// 小于等于
{"type": "value_range", "targets": ["年龄"], "criteria": "<=60"}

// 范围（两个条件组合）
[
  {"type": "value_range", "targets": ["价格"], "criteria": ">=100"},
  {"type": "value_range", "targets": ["价格"], "criteria": "<=500"}
]
```

### Match（集合匹配）

```json
// 单值
{"type": "match", "targets": ["状态"], "criteria": ["完成"]}

// 多值
{"type": "match", "targets": ["部门"], "criteria": ["技术部", "销售部", "财务部"]}

// 多列
{"type": "match", "targets": ["部门", "小组"], "criteria": ["技术部", "前端"]}
```

### Find（字符串查找）

```json
// 包含子串
{"type": "find", "targets": ["姓名"], "criteria": "张"}

// 关键词
{"type": "find", "targets": ["备注"], "criteria": "紧急"}

// 多列查找
{"type": "find", "targets": ["标题", "描述"], "criteria": "重要"}
```

### Equal（列间相等）

```json
// 两列相等
{"type": "equal", "targets": ["预算", "实际"], "criteria": true}

// 三列相等
{"type": "equal", "targets": ["计划", "实际", "预测"], "criteria": true}

// 不相等（有差异）
{"type": "equal", "targets": ["预算", "实际"], "criteria": false}
```

### ExcludeRows（排除行）

```json
// 排除第一行
{"type": "exclude_rows", "criteria": [0, 0]}

// 排除前三行
{"type": "exclude_rows", "criteria": [0, 2]}

// 排除最后两行
{"type": "exclude_rows", "criteria": [-2, -1]}
```

---

## 2. 规则配置模板

### 表头样式

```json
{
  "row_conditions": [{"type": "index", "criteria": [0]}],
  "apply": {"style": "header"}
}
```

### 隔行变色

```json
{
  "row_conditions": [{"type": "index", "criteria": [0, 2, 4, 6, 8]}],
  "apply": {"style": "even_row"}
}
```

### 条件格式（数值）

```json
{
  "row_conditions": [
    {"type": "value_range", "targets": ["销售额"], "criteria": ">50000"}
  ],
  "apply": {
    "style": "normal",
    "overrides": [
      {
        "col_conditions": [
          {"type": "match", "targets": ["销售额"], "criteria": ["销售额"]}
        ],
        "style": "high_value"
      }
    ]
  }
}
```

### 部门合并

```json
{
  "type": "vertical_match",
  "targets": ["部门"]
}
```

### 自动列宽

```json
{
  "target": "column",
  "condition": {"type": "index", "criteria": [0, 1, 2, 3, 4]},
  "value": {"type": "auto"}
}
```

### 固定行高

```json
{
  "target": "row",
  "condition": {"type": "index", "criteria": [0]},
  "value": {"type": "fixed", "value": 30.0}
}
```

---

## 3. 样式属性速查

### 字体

```json
{
  "font_name": "Microsoft YaHei",
  "font_size": 11,
  "font_color": "#000000",
  "bold": true,
  "italic": false,
  "underline": false
}
```

### 背景

```json
{
  "bg_color": "#4472C4"
}
```

### 对齐

```json
{
  "align": "center",
  "valign": "middle",
  "wrap_text": true
}
```

### 边框

```json
{
  "border_left": {"style": "thin", "color": "#000000"},
  "border_right": {"style": "thin", "color": "#000000"},
  "border_top": {"style": "thin", "color": "#000000"},
  "border_bottom": {"style": "thin", "color": "#000000"}
}
```

### 数字格式

```json
{
  "num_format": "#,##0.00"
}
```

常用格式：

| 格式 | 说明 |
|------|------|
| `"0"` | 整数 |
| `"0.00"` | 两位小数 |
| `"#,##0"` | 千分位整数 |
| `"#,##0.00"` | 千分位两位小数 |
| `"0%"` | 百分比 |
| `"yyyy-mm-dd"` | 日期 |

---

## 4. 常见错误速查

### ColumnNotFound

```
Error: ColumnNotFound("销售额")
```

**原因**：列名不存在
**解决**：检查列名拼写，使用 `df.get_column_names()` 查看实际列名

### IndexOutOfBounds

```
Error: IndexOutOfBounds(10, 5)
```

**原因**：索引 10 超出范围（数据只有 5 行）
**解决**：检查数据行数，使用负数索引（如 -1）表示最后一行

### TypeMismatch

```
Error: TypeMismatch("销售额", "String")
```

**原因**：对字符串列使用数值比较
**解决**：确保列的数据类型正确，使用 `df.dtypes()` 查看类型

### JsonError

```
Error: JsonError("missing field `style`")
```

**原因**：JSON 格式错误，缺少必填字段
**解决**：检查 JSON 语法，确保所有必填字段存在

---

## 5. 快速示例

### 完整配置（复制即用）

```json
{
  "style_rules": [
    {
      "row_conditions": [{"type": "index", "criteria": [0]}],
      "apply": {"style": "header"}
    },
    {
      "row_conditions": [
        {"type": "value_range", "targets": ["销售额"], "criteria": ">50000"}
      ],
      "apply": {
        "style": "data",
        "overrides": [
          {
            "col_conditions": [
              {"type": "match", "targets": ["销售额"], "criteria": ["销售额"]}
            ],
            "style": "highlight"
          }
        ]
      }
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
    },
    {
      "target": "column",
      "condition": {"type": "index", "criteria": [0, 1, 2]},
      "value": {"type": "auto"}
    }
  ]
}
```

### Rust 代码示例

```rust
use xlsx_writer::RegionFactory;
use serde_json::json;
use polars::prelude::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 创建 DataFrame
    let df = df! {
        "部门" => ["技术部", "技术部", "销售部"],
        "姓名" => ["张三", "李四", "王五"],
        "销售额" => [60000, 45000, 70000]
    }?;

    // 配置
    let config = json!({
        "style_rules": [
            {
                "row_conditions": [{"type": "index", "criteria": [0]}],
                "apply": {"style": "header"}
            }
        ]
    });

    // 执行
    let factory = RegionFactory::from_json(&config)?;
    let styles = factory.execute(&df)?;

    println!("配置成功！");
    Ok(())
}
```

---

## 相关文档

- [完整配置指南](./configuration_guide.md)
- [基础示例](./examples/basic.json)
- [条件格式示例](./examples/conditional.json)
- [合并单元格示例](./examples/merge.json)
- [行高列宽示例](./examples/dimensions.json)
- [完整示例](./examples/complete.json)
