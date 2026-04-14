# xlsx_writer 配置指南

本文档详细介绍如何使用 JSON 配置来控制 Excel 文件的样式、合并单元格和维度设置。

## 目录

1. [快速开始](#1-快速开始)
2. [配置结构总览](#2-配置结构总览)
3. [StyleRules 详解](#3-stylerules-详解)
4. [MergeRules 详解](#4-mergerules-详解)
5. [DimensionRules 详解](#5-dimensionrules-详解)
6. [条件类型完整参考](#6-条件类型完整参考)
7. [样式定义](#7-样式定义)
8. [常见用例 Cookbook](#8-常见用例-cookbook)
9. [故障排除](#9-故障排除)

---

## 1. 快速开始

### 1.1 最小可用配置

```rust
use xlsx_writer::RegionFactory;
use serde_json::json;

let config = json!({
    "style_rules": [
        {
            "row_conditions": [{"type": "index", "criteria": [0]}],
            "apply": {"style": "header"}
        }
    ]
});

let factory = RegionFactory::from_json(&config)?;
let styles = factory.execute(&df)?;
```

### 1.2 完整配置示例

```json
{
  "style_rules": [
    {
      "row_conditions": [{"type": "index", "criteria": [0]}],
      "apply": {
        "style": "header",
        "overrides": [
          {
            "col_conditions": [{"type": "match", "targets": ["销售额"], "criteria": ["销售额"]}],
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

---

## 2. 配置结构总览

配置文件是一个 JSON 对象，包含四个可选的顶层键：

| 键名 | 类型 | 说明 |
|------|------|------|
| `styles` | 对象 | 定义样式库，包含所有可复用的样式模板 |
| `style_rules` | 数组 | 定义单元格样式的应用规则 |
| `merge_rules` | 数组 | 定义单元格合并规则 |
| `dimension_rules` | 数组 | 定义行高和列宽的设置规则 |

**注意**：所有四个键都是可选的。如果某个键不存在或为空，对应的功能将不会执行。

### 2.1 完整配置结构示例

```json
{
  "styles": {
    "header": {
      "font_name": "Microsoft YaHei",
      "font_size": 12,
      "bold": true,
      "bg_color": "#4472C4",
      "font_color": "#FFFFFF",
      "align": "center"
    },
    "data": {
      "font_size": 10,
      "border": "thin"
    },
    "highlight": {
      "bg_color": "#FFC7CE",
      "font_color": "#9C0006"
    }
  },
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
}
```

### 2.2 各配置键的关系

```
┌─────────────────────────────────────────────────────────────┐
│                     JSON 配置文件                            │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  ┌──────────────┐   定义可复用的样式模板                     │
│  │ styles       │   供 style_rules 引用                      │
│  └──────────────┘                                           │
│         │                                                   │
│         ▼                                                   │
│  ┌──────────────┐   根据条件将样式应用到单元格               │
│  │ style_rules  │   引用 styles 中定义的样式名称             │
│  └──────────────┘                                           │
│                                                             │
│  ┌──────────────┐   定义单元格合并规则                       │
│  │ merge_rules  │                                           │
│  └──────────────┘                                           │
│                                                             │
│  ┌──────────────┐   定义行高和列宽                           │
│  │ dimension_   │                                           │
│  │   rules      │                                           │
│  └──────────────┘                                           │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### 2.3 在代码中使用配置

```rust
use xlsx_writer::Workbook;
use xlsx_writer::RegionFactory;
use serde_json::json;

// 完整的 JSON 配置
let config = json!({
    "styles": {
        "header": { "bg_color": "#4472C4", "bold": true },
        "data": { "font_size": 10 }
    },
    "style_rules": [
        {"row_conditions": [{"type": "index", "criteria": [0]}], "apply": {"style": "header"}}
    ]
});

// 1. 创建 Workbook 并加载样式库
let workbook = Workbook::new()?
    .with_library_from_json(&config["styles"])?;  // 加载样式定义

// 2. 使用 RegionFactory 应用样式规则
let factory = RegionFactory::from_json(&config)?;
let region_styles = factory.execute(&df)?;

// 3. 创建 Region 并添加 Sheet
let region = SheetRegion::from_dataframe(df, "data", Some(true), region_styles)?;
let workbook = workbook
    .add_sheet(WorkSheet::new("报表", vec![region])?)
    .save("output.xlsx")?;
```

**关键点**：
- `styles` 通过 `with_library_from_json()` 方法加载到 Workbook
- `style_rules`、`merge_rules`、`dimension_rules` 通过 `RegionFactory` 执行
- StyleRules 中引用的样式名称必须在 `styles` 中已定义

---

## 3. StyleRules 详解

StyleRules 用于根据数据内容自动应用单元格样式。

### 3.1 基本结构

```json
{
  "row_conditions": [...],  // 选择哪些行应用样式
  "apply": {
    "style": "样式名称",     // 基础样式
    "overrides": [...]       // 可选：列级别的样式覆盖
  }
}
```

### 3.2 规则优先级

- 规则按照数组中的顺序执行
- **后定义的规则优先级更高**，会覆盖前面规则对相同单元格的设置
- 建议将通用规则放在前面，特殊规则放在后面

### 3.3 row_conditions（行条件）

`row_conditions` 是一个条件数组，用于选择满足所有条件的行。

**逻辑关系**：多个条件之间是**逻辑"与"（AND）**关系，只有同时满足所有条件的行才会被选中。

**示例**：

```json
{
  "row_conditions": [
    {"type": "index", "criteria": [0, 1, 2, 3, 4]},
    {"type": "value_range", "targets": ["销售额"], "criteria": ">10000"}
  ],
  "apply": {"style": "high_value"}
}
```

这个规则会选中**索引为 0-4 且销售额大于 10000** 的行。

### 3.4 apply（样式应用）

#### 3.4.1 基础样式（style）

```json
{"style": "header"}
```

- `style` 必须是字符串，对应 StyleLibrary 中预定义的样式名称
- 被选中的所有行的所有单元格首先会应用这个基础样式

#### 3.4.2 覆盖规则（overrides）

`overrides` 允许在已选中的行内，对特定列的单元格应用不同的样式。

```json
{
  "overrides": [
    {
      "col_conditions": [...],  // 列条件
      "style": "覆盖样式"        // 满足条件时应用的样式
    }
  ]
}
```

**示例**：设置表头为蓝色，但"销售额"列的表头为红色

```json
{
  "row_conditions": [{"type": "index", "criteria": [0]}],
  "apply": {
    "style": "blue_header",
    "overrides": [
      {
        "col_conditions": [
          {"type": "match", "targets": ["销售额"], "criteria": ["销售额"]}
        ],
        "style": "red_header"
      }
    ]
  }
}
```

### 3.5 完整示例

**场景**：销售数据表格，要求：
1. 第一行（表头）使用蓝色背景
2. 销售额大于 50000 的行使用浅红色背景
3. 销售额大于 50000 的单元格本身使用深红色字体

```json
{
  "style_rules": [
    {
      "row_conditions": [{"type": "index", "criteria": [0]}],
      "apply": {"style": "blue_header"}
    },
    {
      "row_conditions": [
        {"type": "value_range", "targets": ["销售额"], "criteria": ">50000"}
      ],
      "apply": {
        "style": "light_red_row",
        "overrides": [
          {
            "col_conditions": [
              {"type": "match", "targets": ["销售额"], "criteria": ["销售额"]}
            ],
            "style": "dark_red_cell"
          }
        ]
      }
    }
  ]
}
```

---

## 4. MergeRules 详解

MergeRules 用于自动合并具有相同值的相邻单元格。

### 4.1 基本结构

```json
{
  "type": "vertical_match",
  "targets": ["列名1", "列名2"]
}
```

### 4.2 规则类型

#### 4.2.1 vertical_match（纵向合并）

在指定列中，将连续相同的值纵向合并。

```json
{
  "type": "vertical_match",
  "targets": ["部门"]
}
```

**效果示例**：

| 部门 | 姓名 |
|------|------|
| 技术部 | 张三 |
| 技术部 | 李四 |
| 销售部 | 王五 |

合并后，"技术部"单元格会跨两行显示。

#### 4.2.2 horizontal_match（横向合并）

在指定行中，将连续相同的值横向合并。（暂未实现）

### 4.3 多列合并

可以同时指定多个列进行合并：

```json
{
  "type": "vertical_match",
  "targets": ["部门", "小组"]
}
```

**注意**：合并的顺序遵循 `targets` 数组的顺序。

### 4.4 完整示例

```json
{
  "merge_rules": [
    {"type": "vertical_match", "targets": ["部门"]},
    {"type": "vertical_match", "targets": ["小组"]}
  ]
}
```

---

## 5. DimensionRules 详解

DimensionRules 用于设置行高和列宽。

### 5.1 基本结构

```json
{
  "target": "row" | "column",
  "condition": {...},
  "value": {"type": "fixed", "value": 30.0} | {"type": "auto"}
}
```

### 5.2 target（目标维度）

| 值 | 说明 |
|----|------|
| `"row"` | 设置行高 |
| `"column"` | 设置列宽 |

### 5.3 condition（条件）

DimensionRules 复用 StyleRules 的条件系统。可用的条件类型：

- **Row（行高）**：所有 6 种条件类型都可用（Index, ValueRange, Match, Find, Equal, ExcludeRows）
- **Column（列宽）**：仅支持 Index、Match、Find；ValueRange 和 Equal 会被静默忽略

### 5.4 value（值）

#### 5.4.1 Fixed（固定值）

```json
{"type": "fixed", "value": 30.0}
```

- 对于行高：单位是磅（point）
- 对于列宽：单位是字符宽度

#### 5.4.2 Auto（自动计算）

```json
{"type": "auto"}
```

- **列宽**：根据内容自动计算
  - 中文字符（Unicode 0x4E00-0x9FFF）计 2 宽度单位
  - 其他字符计 1 宽度单位
  - 结果加上 2 单位的内边距
  - 最小宽度为 8
- **行高**：暂不支持（TODO），使用默认值 15

### 5.5 完整示例

**场景**：设置表头行高为 30，自动调整前三列的列宽

```json
{
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

### 5.6 列宽计算的注意事项

1. **多列分别计算**：每列的 Auto 宽度是根据该列自身内容独立计算的
2. **最小宽度限制**：即使内容很短，列宽也不会小于 8
3. **中文字符处理**：会自动识别中文字符并分配更大的宽度

---

## 6. 条件类型完整参考

本章详细说明所有可用的条件类型。

### 6.1 Index（索引定位）

通过物理行号或列号直接选择目标。

**语法**：

```json
{
  "type": "index",
  "targets": [],      // 通常为空
  "criteria": [0, 1, -1]  // 索引列表
}
```

**参数说明**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `targets` | 字符串数组 | 否 | 通常为空数组或省略 |
| `criteria` | 整数数组 | 是 | 索引列表，支持负数 |

**索引规则**：

- **正数索引**：0 表示第一行数据（不包括表头）
- **负数索引**：-1 表示最后一行数据，-2 表示倒数第二行，以此类推
- **数组**：可以同时指定多个索引

**示例**：

```json
// 选择第一行
{"type": "index", "criteria": [0]}

// 选择前三行
{"type": "index", "criteria": [0, 1, 2]}

// 选择最后一行
{"type": "index", "criteria": [-1]}

// 选择第一行和最后一行
{"type": "index", "criteria": [0, -1]}
```

**边界检查**：
- 正数索引必须小于数据行数
- 负数索引转换后必须大于等于 0
- 越界会导致错误

### 6.2 ValueRange（数值范围）

检查指定列的数值是否满足比较表达式。

**语法**：

```json
{
  "type": "value_range",
  "targets": ["列名"],
  "criteria": ">10000"
}
```

**参数说明**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `targets` | 字符串数组 | 是 | 要检查的列名列表 |
| `criteria` | 字符串 | 是 | 比较表达式 |

**支持的比较运算符**：

| 运算符 | 含义 | 示例 |
|--------|------|------|
| `>` | 大于 | `">100"` |
| `<` | 小于 | `"<100"` |
| `>=` | 大于等于 | `">=0.5"` |
| `<=` | 小于等于 | `"<=50"` |
| `=` | 等于（默认） | `"=100"` 或 `"100"` |

**示例**：

```json
// 销售额大于 10000
{"type": "value_range", "targets": ["销售额"], "criteria": ">10000"}

// 年龄介于 18 到 60 之间（使用两个条件）
[
  {"type": "value_range", "targets": ["年龄"], "criteria": ">=18"},
  {"type": "value_range", "targets": ["年龄"], "criteria": "<=60"}
]

// 多列同时检查
{"type": "value_range", "targets": ["销售额", "利润"], "criteria": ">0"}
```

**注意事项**：
- 只能用于数值类型的列
- 如果列包含非数值数据，会报错
- 在列宽规则中使用会被静默忽略

### 6.3 Match（集合匹配）

检查指定列的值是否在给定的白名单中。

**语法**：

```json
{
  "type": "match",
  "targets": ["列名"],
  "criteria": ["值1", "值2", "值3"]
}
```

**参数说明**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `targets` | 字符串数组 | 是 | 要检查的列名列表 |
| `criteria` | 字符串数组 | 是 | 允许的值列表 |

**匹配规则**：

- 字符串列：精确匹配字符串值
- 数值列：将数值转为字符串后匹配
- 空值（null）不会被匹配

**示例**：

```json
// 部门是"技术部"或"销售部"
{"type": "match", "targets": ["部门"], "criteria": ["技术部", "销售部"]}

// 状态是"完成"或"进行中"
{"type": "match", "targets": ["状态"], "criteria": ["完成", "进行中"]}

// 多列匹配
{"type": "match", "targets": ["部门", "小组"], "criteria": ["技术部", "前端"]}
```

**在列宽规则中的使用**：

可以用于根据列名选择列：

```json
{
  "target": "column",
  "condition": {
    "type": "match",
    "targets": ["user_name", "user_age"],  // 在 DataFrame 的列名中查找
    "criteria": ["user_name", "user_age"]   // 匹配这些列名
  },
  "value": {"type": "fixed", "value": 20.0}
}
```

### 6.4 Find（字符串查找）

检查指定字符串列是否包含特定子串。

**语法**：

```json
{
  "type": "find",
  "targets": ["列名"],
  "criteria": "子串"
}
```

**参数说明**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `targets` | 字符串数组 | 是 | 要检查的列名列表（必须是字符串类型） |
| `criteria` | 字符串 | 是 | 要查找的子串 |

**示例**：

```json
// 姓名包含"张"
{"type": "find", "targets": ["姓名"], "criteria": "张"}

// 备注包含"紧急"
{"type": "find", "targets": ["备注"], "criteria": "紧急"}
```

**注意事项**：
- 只能用于字符串类型的列
- 区分大小写
- 不支持正则表达式（简单子串匹配）

**在列宽规则中的使用**：

可以用于匹配包含特定关键词的列名：

```json
{
  "target": "column",
  "condition": {
    "type": "find",
    "targets": ["date", "time", "name"],  // 在 DataFrame 的列名中查找
    "criteria": "date"                      // 匹配包含"date"的列名
  },
  "value": {"type": "auto"}
}
```

### 6.5 Equal（列间相等）

比较多个列的值是否相等。

**语法**：

```json
{
  "type": "equal",
  "targets": ["列名1", "列名2", "列名3"],
  "criteria": true
}
```

**参数说明**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `targets` | 字符串数组 | 是 | 要比较的列名列表（至少 2 个） |
| `criteria` | 布尔值 | 是 | `true` 表示要求全部相等，`false` 表示要求存在不相等 |

**示例**：

```json
// 预算和实际支出相等
{"type": "equal", "targets": ["预算", "实际支出"], "criteria": true}

// 存在差异（预算不等于实际支出）
{"type": "equal", "targets": ["预算", "实际支出"], "criteria": false}
```

**注意事项**：
- 至少需要指定 2 个列
- 类型不兼容的比较会报错
- 在列宽规则中使用会被静默忽略

### 6.6 ExcludeRows（排除行）

排除特定范围内的行。

**语法**：

```json
{
  "type": "exclude_rows",
  "targets": [],
  "criteria": [0, 0]
}
```

**参数说明**：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `targets` | 字符串数组 | 否 | 通常为空数组 |
| `criteria` | 整数数组（2个元素） | 是 | `[开始, 结束]`，包含起止 |

**索引规则**：

与 Index 条件相同，支持正数和负数索引。

**示例**：

```json
// 排除第一行数据
{"type": "exclude_rows", "criteria": [0, 0]}

// 排除最后两行
{"type": "exclude_rows", "criteria": [-2, -1]}

// 排除第 2 到第 5 行
{"type": "exclude_rows", "criteria": [1, 4]}
```

**与其他条件的关系**：

ExcludeRows 是**逻辑"非"**操作。如果某行被 ExcludeRows 排除，即使满足其他所有条件，也不会被选中。

**示例**：选中高销售额但排除第一行

```json
{
  "row_conditions": [
    {"type": "value_range", "targets": ["销售额"], "criteria": ">10000"},
    {"type": "exclude_rows", "criteria": [0, 0]}
  ]
}
```

### 6.7 条件类型对比表

| 条件类型 | 用途 | 数据驱动 | 列宽规则支持 |
|---------|------|----------|-------------|
| `Index` | 物理位置 | 否 | ✅ |
| `ValueRange` | 数值范围 | 是 | ❌（静默忽略） |
| `Match` | 集合匹配 | 是 | ✅（匹配列名） |
| `Find` | 子串查找 | 是 | ✅（匹配列名） |
| `Equal` | 列间比较 | 是 | ❌（静默忽略） |
| `ExcludeRows` | 排除范围 | 否 | ❌（无意义） |

---

## 7. 样式定义

StyleRules 中使用的样式名称必须在 StyleLibrary 中预定义。样式可以通过 JSON 配置文件定义，然后在代码中加载。

### 7.1 通过 JSON 定义样式

样式定义位于 JSON 配置的 `styles` 键下，是一个对象，键为样式名称，值为样式属性。

**基本结构**：

```json
{
  "styles": {
    "样式名称1": {
      "属性1": "值1",
      "属性2": "值2"
    },
    "样式名称2": {
      ...
    }
  }
}
```

**完整示例**：

```json
{
  "styles": {
    "header": {
      "font_name": "Microsoft YaHei",
      "font_size": 12,
      "bold": true,
      "bg_color": "#4472C4",
      "font_color": "#FFFFFF",
      "align": "center",
      "valign": "middle"
    },
    "data": {
      "font_size": 10,
      "border": "thin"
    },
    "highlight": {
      "bg_color": "#FFC7CE",
      "font_color": "#9C0006",
      "bold": true
    },
    "number": {
      "num_format": "#,##0.00",
      "align": "right"
    },
    "date": {
      "num_format": "yyyy-mm-dd",
      "align": "center"
    }
  },
  "style_rules": [
    {
      "row_conditions": [{"type": "index", "criteria": [0]}],
      "apply": {"style": "header"}
    },
    {
      "row_conditions": [{"type": "index", "criteria": [1, 2, 3, 4, 5]}],
      "apply": {"style": "data"}
    }
  ]
}
```

### 7.2 在代码中加载样式

#### 7.2.1 从 JSON 配置加载

```rust
use xlsx_writer::Workbook;
use serde_json::json;

// 完整的配置（包含样式和规则）
let config = json!({
    "styles": {
        "header": { "bg_color": "#4472C4", "bold": true },
        "data": { "font_size": 10 }
    },
    "style_rules": [
        {"row_conditions": [{"type": "index", "criteria": [0]}], "apply": {"style": "header"}}
    ]
});

// 创建 Workbook 并加载样式库
let workbook = Workbook::new()?
    .with_library_from_json(&config["styles"])?;  // 加载 styles 部分
```

#### 7.2.2 单独加载样式配置

```rust
// 只包含样式的配置
let style_config = json!({
    "header": { "bg_color": "#4472C4", "bold": true },
    "data": { "font_size": 10 }
});

let workbook = Workbook::new()?
    .with_library_from_json(&style_config)?;
```

#### 7.2.3 从文件加载

```rust
use std::fs;

// 读取配置文件
let config_str = fs::read_to_string("config.json")?;
let config: serde_json::Value = serde_json::from_str(&config_str)?;

// 加载样式
let workbook = Workbook::new()?
    .with_library_from_json(&config["styles"])?;

// 使用 RegionFactory 执行规则
let factory = RegionFactory::from_json(&config)?;
let styles = factory.execute(&df)?;
```

### 7.3 样式属性完整参考

#### 7.3.1 字体属性

| 属性名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `font_name` | 字符串 | `"Calibri"` | 字体名称，如 `"Microsoft YaHei"` |
| `font_size` | 数字 | 11 | 字体大小（磅） |
| `font_color` | 字符串 | `"#000000"` | 字体颜色（十六进制，如 `"#FF0000"`） |
| `bold` | 布尔值 | false | 是否粗体 |
| `italic` | 布尔值 | false | 是否斜体 |
| `underline` | 字符串 | - | 下划线类型：`"single"`, `"double"` |

**示例**：

```json
{
  "title": {
    "font_name": "Microsoft YaHei",
    "font_size": 16,
    "bold": true,
    "font_color": "#FFFFFF"
  }
}
```

#### 7.3.2 对齐属性

| 属性名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `align` | 字符串 | `"left"` | 水平对齐：`"left"`, `"center"`, `"right"`, `"fill"`, `"justify"` |
| `valign` | 字符串 | `"bottom"` | 垂直对齐：`"top"`, `"vcenter"`/`"center"`, `"bottom"`, `"vjustify"` |
| `text_wrap` | 布尔值 | false | 是否自动换行 |

**示例**：

```json
{
  "header": {
    "align": "center",
    "valign": "middle",
    "text_wrap": true
  }
}
```

#### 7.3.3 背景属性

| 属性名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `bg_color` | 字符串 | - | 背景颜色（十六进制） |
| `fg_color` | 字符串 | - | 前景颜色（十六进制） |
| `pattern` | 数字 | - | 填充模式：0（无）, 1（纯色）, 其他模式 |

**示例**：

```json
{
  "highlight": {
    "bg_color": "#FFC7CE",
    "pattern": 1
  }
}
```

#### 7.3.4 边框属性

支持两种方式设置边框：

**方式 1：统一设置（所有边框相同）**

| 属性名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `border` | 字符串 | - | 边框样式：`"thin"`, `"medium"`, `"thick"`, `"dashed"`, `"double"` |
| `border_color` | 字符串 | - | 边框颜色（十六进制） |

**方式 2：分别设置（每边不同）**

| 属性名 | 类型 | 说明 |
|--------|------|------|
| `border_left` | 对象 | 左边框样式 |
| `border_right` | 对象 | 右边框样式 |
| `border_top` | 对象 | 上边框样式 |
| `border_bottom` | 对象 | 下边框样式 |

边框样式对象结构：

```json
{
  "style": "thin",        // 线条样式
  "color": "#000000"      // 线条颜色
}
```

**示例**：

```json
{
  "table_header": {
    "border": "medium",
    "border_color": "#000000"
  },
  "table_cell": {
    "border_left": {"style": "thin", "color": "#CCCCCC"},
    "border_right": {"style": "thin", "color": "#CCCCCC"},
    "border_top": {"style": "thin", "color": "#CCCCCC"},
    "border_bottom": {"style": "thin", "color": "#CCCCCC"}
  }
}
```

#### 7.3.5 数字格式

| 属性名 | 类型 | 默认值 | 说明 |
|--------|------|------|------|
| `num_format` | 字符串 | `"General"` | 数字格式字符串 |

**常用格式**：

| 格式字符串 | 示例输出 | 说明 |
|-----------|---------|------|
| `"General"` | 1234.5 | 通用格式 |
| `"0"` | 1235 | 整数（四舍五入） |
| `"0.00"` | 1234.50 | 保留两位小数 |
| `"#,##0"` | 1,235 | 千分位整数 |
| `"#,##0.00"` | 1,234.50 | 千分位两位小数 |
| `"0%"` | 123450% | 百分比 |
| `"0.00%"` | 123450.00% | 百分比两位小数 |
| `"¥#,##0.00"` | ¥1,234.50 | 人民币格式 |
| `"yyyy-mm-dd"` | 2024-01-15 | 日期格式 |
| `"yyyy-mm-dd hh:mm:ss"` | 2024-01-15 09:30:00 | 日期时间格式 |

**示例**：

```json
{
  "currency": {
    "num_format": "¥#,##0.00",
    "align": "right"
  },
  "percentage": {
    "num_format": "0.00%",
    "align": "center"
  },
  "date": {
    "num_format": "yyyy-mm-dd",
    "align": "center"
  }
}
```

### 7.4 样式继承和覆盖

样式规则中的 `apply` 和 `overrides` 可以引用在 `styles` 中定义的任何样式名称。

**样式引用关系**：

```
styles (JSON配置)
  ├── "header" ───────────────┐
  ├── "data"                  │
  ├── "highlight"             │
  └── "number"                │
                              ▼
style_rules (JSON配置)    被引用
  └── apply.style = "header" ─┘
      └── overrides
          └── style = "highlight"
```

**重要提示**：
- `styles` 中定义的样式名称是全局的，可以在多个 `style_rules` 中重复使用
- 样式名称区分大小写
- 如果在 `style_rules` 中引用了未在 `styles` 中定义的样式名称，会导致错误

### 7.5 完整样式库示例

```json
{
  "styles": {
    "title": {
      "font_name": "Microsoft YaHei",
      "font_size": 16,
      "bold": true,
      "bg_color": "#4472C4",
      "font_color": "#FFFFFF",
      "align": "center",
      "valign": "middle"
    },
    "header": {
      "font_name": "Microsoft YaHei",
      "font_size": 11,
      "bold": true,
      "bg_color": "#D9E1F2",
      "align": "center",
      "valign": "middle",
      "border": "thin",
      "border_color": "#000000"
    },
    "data": {
      "font_name": "Microsoft YaHei",
      "font_size": 10,
      "border_left": {"style": "thin", "color": "#CCCCCC"},
      "border_right": {"style": "thin", "color": "#CCCCCC"},
      "border_top": {"style": "thin", "color": "#CCCCCC"},
      "border_bottom": {"style": "thin", "color": "#CCCCCC"}
    },
    "data_alt": {
      "font_name": "Microsoft YaHei",
      "font_size": 10,
      "bg_color": "#F2F2F2",
      "border_left": {"style": "thin", "color": "#CCCCCC"},
      "border_right": {"style": "thin", "color": "#CCCCCC"},
      "border_top": {"style": "thin", "color": "#CCCCCC"},
      "border_bottom": {"style": "thin", "color": "#CCCCCC"}
    },
    "number": {
      "num_format": "#,##0.00",
      "align": "right"
    },
    "currency": {
      "num_format": "¥#,##0.00",
      "align": "right",
      "font_color": "#C65911"
    },
    "date": {
      "num_format": "yyyy-mm-dd",
      "align": "center"
    },
    "high_value": {
      "bg_color": "#FFC7CE",
      "font_color": "#9C0006",
      "bold": true
    },
    "low_value": {
      "bg_color": "#C6EFCE",
      "font_color": "#006100"
    },
    "warning": {
      "bg_color": "#FFEB9C",
      "font_color": "#9C5700"
    },
    "summary": {
      "font_name": "Microsoft YaHei",
      "font_size": 11,
      "bold": true,
      "bg_color": "#FFF2CC",
      "border_top": {"style": "medium", "color": "#000000"}
    },
    "note": {
      "font_size": 9,
      "italic": true,
      "font_color": "#666666"
    }
  },
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
            "style": "high_value"
          }
        ]
      }
    }
  ]
}
```

### 7.2 样式属性参考

#### 7.2.1 字体属性

| 属性名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `font_name` | 字符串 | `"Calibri"` | 字体名称 |
| `font_size` | 数字 | 11 | 字体大小（磅） |
| `font_color` | 字符串 | `"#000000"` | 字体颜色（十六进制） |
| `bold` | 布尔值 | false | 是否粗体 |
| `italic` | 布尔值 | false | 是否斜体 |
| `underline` | 布尔值 | false | 是否下划线 |

#### 7.2.2 背景属性

| 属性名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `bg_color` | 字符串 | - | 背景颜色（十六进制） |

#### 7.2.3 对齐属性

| 属性名 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `align` | 字符串 | `"left"` | 水平对齐：`left`, `center`, `right` |
| `valign` | 字符串 | `"bottom"` | 垂直对齐：`top`, `middle`, `bottom` |
| `wrap_text` | 布尔值 | false | 是否自动换行 |

#### 7.2.4 边框属性

| 属性名 | 类型 | 默认值 | 说明 |
|--------|------|------|------|
| `border_left` | 对象 | - | 左边框样式 |
| `border_right` | 对象 | - | 右边框样式 |
| `border_top` | 对象 | - | 上边框样式 |
| `border_bottom` | 对象 | - | 下边框样式 |

边框样式对象结构：

```json
{
  "style": "thin",        // 线条样式：thin, medium, thick
  "color": "#000000"      // 线条颜色
}
```

#### 7.2.5 数字格式

| 属性名 | 类型 | 默认值 | 说明 |
|--------|------|------|------|
| `num_format` | 字符串 | `"General"` | 数字格式 |

常用数字格式：

| 格式 | 说明 |
|------|------|
| `"General"` | 通用格式 |
| `"0"` | 整数 |
| `"0.00"` | 保留两位小数 |
| `"#,##0"` | 千分位整数 |
| `"#,##0.00"` | 千分位两位小数 |
| `"0%"` | 百分比 |
| `"yyyy-mm-dd"` | 日期格式 |
| `"yyyy-mm-dd hh:mm:ss"` | 日期时间格式 |

### 7.3 完整样式示例

```json
{
  "title": {
    "font_name": "Microsoft YaHei",
    "font_size": 16,
    "bold": true,
    "bg_color": "#4472C4",
    "font_color": "#FFFFFF",
    "align": "center",
    "valign": "middle"
  },
  "header": {
    "font_name": "Microsoft YaHei",
    "font_size": 11,
    "bold": true,
    "bg_color": "#D9E1F2",
    "align": "center",
    "valign": "middle",
    "border_left": {"style": "thin", "color": "#000000"},
    "border_right": {"style": "thin", "color": "#000000"},
    "border_top": {"style": "thin", "color": "#000000"},
    "border_bottom": {"style": "thin", "color": "#000000"}
  },
  "data": {
    "font_name": "Microsoft YaHei",
    "font_size": 10,
    "border_left": {"style": "thin", "color": "#CCCCCC"},
    "border_right": {"style": "thin", "color": "#CCCCCC"}
  },
  "number": {
    "num_format": "#,##0.00",
    "align": "right"
  },
  "date": {
    "num_format": "yyyy-mm-dd",
    "align": "center"
  },
  "high_value": {
    "bg_color": "#FFC7CE",
    "font_color": "#9C0006",
    "bold": true
  },
  "low_value": {
    "bg_color": "#C6EFCE",
    "font_color": "#006100"
  }
}
```

---

## 8. 常见用例 Cookbook

本章提供常见场景的完整配置示例。

### 8.1 基础表格样式

**需求**：
- 表头蓝色背景，白色粗体字
- 数据区域边框
- 数值右对齐

```json
{
  "style_rules": [
    {
      "row_conditions": [{"type": "index", "criteria": [0]}],
      "apply": {"style": "header"}
    },
    {
      "row_conditions": [{"type": "index", "criteria": [1, 2, 3, 4, 5]}],
      "apply": {
        "style": "data",
        "overrides": [
          {
            "col_conditions": [
              {"type": "match", "targets": ["销售额", "利润"], "criteria": ["销售额", "利润"]}
            ],
            "style": "number"
          }
        ]
      }
    }
  ]
}
```

### 8.2 条件格式（数据条）

**需求**：销售额大于 50000 的单元格显示为红色

```json
{
  "style_rules": [
    {
      "row_conditions": [
        {"type": "value_range", "targets": ["销售额"], "criteria": ">50000"}
      ],
      "apply": {
        "style": "data_row",
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
  ]
}
```

### 8.3 隔行变色

**需求**：奇数行和偶数行不同背景色

```json
{
  "style_rules": [
    {
      "row_conditions": [{"type": "index", "criteria": [0, 2, 4, 6, 8]}],
      "apply": {"style": "even_row"}
    },
    {
      "row_conditions": [{"type": "index", "criteria": [1, 3, 5, 7, 9]}],
      "apply": {"style": "odd_row"}
    }
  ]
}
```

### 8.4 部门合并

**需求**：相同部门的单元格纵向合并

```json
{
  "merge_rules": [
    {"type": "vertical_match", "targets": ["部门"]}
  ]
}
```

### 8.5 多条件样式

**需求**：
- 技术部的员工使用蓝色背景
- 销售部的员工且销售额大于 50000 使用红色背景

```json
{
  "style_rules": [
    {
      "row_conditions": [
        {"type": "match", "targets": ["部门"], "criteria": ["技术部"]}
      ],
      "apply": {"style": "tech_dept"}
    },
    {
      "row_conditions": [
        {"type": "match", "targets": ["部门"], "criteria": ["销售部"]},
        {"type": "value_range", "targets": ["销售额"], "criteria": ">50000"}
      ],
      "apply": {"style": "sales_high"}
    }
  ]
}
```

### 8.6 复杂表格布局

**需求**：
- 表头固定行高 30
- 自动调整所有列宽
- 合并部门列
- 销售额条件格式

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
            "style": "high_value"
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
      "condition": {"type": "index", "criteria": [0, 1, 2, 3, 4]},
      "value": {"type": "auto"}
    }
  ]
}
```

### 8.7 分组统计表

**需求**：
- 一级分组：部门（合并）
- 二级分组：小组（合并）
- 汇总行粗体显示

```json
{
  "style_rules": [
    {
      "row_conditions": [{"type": "index", "criteria": [0]}],
      "apply": {"style": "header"}
    },
    {
      "row_conditions": [
        {"type": "find", "targets": ["姓名"], "criteria": "合计"}
      ],
      "apply": {"style": "summary"}
    }
  ],
  "merge_rules": [
    {"type": "vertical_match", "targets": ["部门"]},
    {"type": "vertical_match", "targets": ["小组"]}
  ]
}
```

### 8.8 财务报表

**需求**：
- 预算和实际支出相等的行标记为绿色
- 存在差异的行标记为黄色
- 货币格式

```json
{
  "style_rules": [
    {
      "row_conditions": [{"type": "index", "criteria": [0]}],
      "apply": {"style": "header"}
    },
    {
      "row_conditions": [
        {"type": "equal", "targets": ["预算", "实际支出"], "criteria": true}
      ],
      "apply": {"style": "balanced"}
    },
    {
      "row_conditions": [
        {"type": "equal", "targets": ["预算", "实际支出"], "criteria": false}
      ],
      "apply": {"style": "unbalanced"}
    }
  ]
}
```

### 8.9 多Sheet配置

**需求**：不同的 Sheet 使用不同的配置

```rust
// Sheet 1: 销售数据
let sales_config = json!({
    "style_rules": [...],
    "merge_rules": [...]
});

// Sheet 2: 部门统计
let dept_config = json!({
    "style_rules": [...],
    "dimension_rules": [...]
});

let workbook = Workbook::new()?
    .with_dataframe(sales_df, &sales_config)?
    .with_dataframe(dept_df, &dept_config)?;
```

### 8.10 排除特定行

**需求**：对高价值订单应用样式，但排除第一行（可能是标题行）

```json
{
  "style_rules": [
    {
      "row_conditions": [
        {"type": "value_range", "targets": ["金额"], "criteria": ">100000"},
        {"type": "exclude_rows", "criteria": [0, 0]}
      ],
      "apply": {"style": "vip_order"}
    }
  ]
}
```

---

## 9. 故障排除

### 9.1 常见问题

#### Q1: 样式没有生效

**可能原因**：
1. 样式名称拼写错误
2. 样式未在 StyleLibrary 中定义
3. 条件没有匹配到任何行

**排查步骤**：
1. 检查样式名称大小写
2. 确认样式已正确加载
3. 添加调试打印检查数据是否符合条件

#### Q2: 列宽 Auto 计算结果不正确

**可能原因**：
1. 内容宽度小于最小宽度（8）
2. 中文字符识别失败

**解决方案**：
1. 检查数据内容长度
2. 确认使用中文字符（Unicode 0x4E00-0x9FFF）

#### Q3: 合并单元格没有生效

**可能原因**：
1. 合并规则配置错误
2. 相邻单元格的值不相同（包括空格、大小写等）

**排查步骤**：
1. 检查合并规则的 `targets` 是否正确
2. 打印数据确认值完全相同

#### Q4: 负数索引报错

**可能原因**：
1. 数据行数不足
2. 负数索引转换后越界

**示例**：
- DataFrame 只有 2 行数据
- 使用 `{"type": "index", "criteria": [-3]}`
- -3 转换后 = 2 + (-3) + 1 = 0（有效）
- 但使用 -4 转换后 = -1（越界报错）

#### Q5: 数值比较报错 "TypeMismatch"

**可能原因**：
列的数据类型不是数值类型

**解决方案**：
1. 检查 DataFrame 的列类型
2. 确保数据已正确解析为数值

### 9.2 调试技巧

#### 技巧 1: 打印 RegionStyles

```rust
let styles = factory.execute(&df)?;
println!("Cell styles: {:?}", styles.cell_styles);
println!("Merge ranges: {:?}", styles.merge_ranges);
println!("Row heights: {:?}", styles.row_heights);
println!("Col widths: {:?}", styles.col_widths);
```

#### 技巧 2: 分步测试配置

先测试简单的配置，逐步添加复杂度：

```rust
// 步骤 1: 只测试样式
let config = json!({
    "style_rules": [...]
});

// 步骤 2: 添加合并
let config = json!({
    "style_rules": [...],
    "merge_rules": [...]
});

// 步骤 3: 添加维度
let config = json!({
    "style_rules": [...],
    "merge_rules": [...],
    "dimension_rules": [...]
});
```

#### 技巧 3: 检查 DataFrame 内容

```rust
println!("DataFrame:\n{:?}", df);
println!("Columns: {:?}", df.get_column_names());
println!("Height: {}", df.height());
```

### 9.3 性能优化建议

1. **减少规则数量**：过多的规则会影响性能，尽量合并相似规则
2. **精确条件**：使用更精确的条件减少匹配范围
3. **避免复杂条件**：复杂的条件组合会增加计算时间
4. **预处理数据**：在传入 DataFrame 前处理好数据类型

### 9.4 错误代码速查

| 错误 | 说明 | 解决方案 |
|------|------|----------|
| `ColumnNotFound` | 列名不存在 | 检查列名拼写 |
| `IndexOutOfBounds` | 索引越界 | 检查索引范围 |
| `TypeMismatch` | 类型不匹配 | 检查数据类型 |
| `JsonError` | JSON 解析错误 | 检查 JSON 语法 |

---

## 附录 A: 版本历史

- **v0.1.0**: 初始版本，支持基础样式规则和合并规则
- **v0.2.0**: 新增 DimensionRules，支持行高列宽设置
- **v0.3.0**: 支持更多条件类型，优化性能

## 附录 B: 相关资源

- [项目仓库](https://github.com/your-org/xlsx_writer)
- [API 文档](https://docs.rs/xlsx_writer)
- [示例代码](../examples/)
