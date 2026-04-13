# xlsx_writer

一个专为数据分析自动化设计的 Rust Excel 处理库，支持 **JSON 配置驱动**的样式和合并规则，同时兼容 **.xlsx** 和 **.xls** 双格式。

## ✨ 主要特性

- **📝 JSON 配置驱动**：通过 JSON 文件定义样式规则和合并区域，无需修改代码即可调整输出格式
- **📊 Polars 原生集成**：与 Polars DataFrame 无缝协作，支持大规模数据处理
- **🎨 智能样式系统**：基于数据内容的条件样式（数值范围、字符串匹配、列间比较等）
- **🔗 自动合并区域**：支持纵向、横向自动合并，以及静态坐标合并
- **💾 双格式支持**：同时支持现代 .xlsx 格式和旧版 .xls (Excel 97-2003) 格式
- **🖥️ 跨平台字体**：自动适配 Windows（微软雅黑）和 macOS（苹方）的中文字体

## 🚀 快速开始

### 安装

在 `Cargo.toml` 中添加依赖：

```toml
[dependencies]
xlsx_writer = { git = "https://github.com/yourusername/xlsx_writer" }
polars = { version = "0.53.0", features = ["lazy", "dtype-full"] }
serde_json = "1.0"
```

### 基础示例

```rust
use xlsx_writer::prelude::*;
use polars::prelude::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 1. 准备数据
    let df = df!(
        "部门" => &["技术部", "技术部", "财务部", "财务部"],
        "姓名" => &["张三", "李四", "王五", "赵六"],
        "薪资" => &[15000.0, 18000.0, 12000.0, 14000.0]
    )?;

    // 2. 创建并保存工作簿
    Workbook::new()?
        .insert(df, Some("工资表".into()), None, None)?
        .save("output.xlsx")?;

    Ok(())
}
```

---

## 📖 使用指南

### 1. JSON 配置驱动（推荐）

通过 JSON 配置文件定义样式和合并规则，实现**零代码**格式调整：

```rust
use xlsx_writer::prelude::*;
use polars::prelude::*;
use serde_json::json;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let df = df!(
        "部门" => &["技术部", "技术部", "财务部", "财务部"],
        "姓名" => &["张三", "李四", "王五", "赵六"],
        "薪资" => &[15000.0, 18000.0, 12000.0, 14000.0]
    )?;

    // JSON 配置：定义样式规则和合并规则
    let config = json!({
        // 样式规则：薪资大于 15000 的单元格标红
        "style_rules": [
            {
                "row_conditions": [
                    {"type": "value_range", "targets": ["薪资"], "criteria": ">15000"}
                ],
                "apply": {
                    "style": "highlight_red",
                    "overrides": [
                        {"style": "highlight_red", "col_conditions": [{"type": "match", "targets": ["薪资"], "criteria": ["薪资"]}]}
                    ]
                }
            }
        ],
        // 合并规则：纵向合并相同部门
        "merge_rules": [
            {"type": "vertical_match", "targets": ["部门"]}
        ]
    });

    // 定义样式库
    let library = json!({
        "highlight_red": {
            "font_color": "#FF0000",
            "bold": true
        }
    });

    Workbook::new()?
        .with_library_from_json(&library)?
        .insert_with_config(df, Some("工资表".into()), &config)?
        .save("output.xlsx")?;

    Ok(())
}
```

---

### 2. 样式规则详解

样式规则支持多种条件类型，可以精确控制单元格样式：

#### 条件类型

| 类型 | 说明 | 示例 |
|------|------|------|
| `index` | 按行索引定位 | `{"type": "index", "criteria": [0, 1, -1]}`（第1、2行和最后一行） |
| `value_range` | 数值范围比较 | `{"type": "value_range", "targets": ["薪资"], "criteria": ">=10000"}` |
| `match` | 字符串匹配 | `{"type": "match", "targets": ["部门"], "criteria": ["技术部", "财务部"]}` |
| `find` | 子串查找 | `{"type": "find", "targets": ["姓名"], "criteria": "张"}` |
| `equal` | 列间相等 | `{"type": "equal", "targets": ["计划", "实际"], "criteria": true}` |
| `exclude_rows` | 排除行范围 | `{"type": "exclude_rows", "criteria": [0, 0]}`（排除第一行） |

#### 完整样式规则示例

```json
{
  "style_rules": [
    {
      "row_conditions": [
        {"type": "value_range", "targets": ["销售额"], "criteria": ">100000"}
      ],
      "apply": {
        "style": "high_performer",
        "overrides": [
          {
            "style": "money_format",
            "col_conditions": [{"type": "match", "targets": ["销售额"], "criteria": ["销售额"]}]
          }
        ]
      }
    }
  ]
}
```

---

### 3. 合并规则详解

自动计算并生成 Excel 合并区域：

#### 合并类型

| 类型 | 说明 | 示例 |
|------|------|------|
| `fixed` | 静态坐标 | `{"type": "fixed", "targets": [[0, 0, 0, 3]]}`（合并 A1:D1） |
| `vertical_match` | 纵向合并 | `{"type": "vertical_match", "targets": ["部门", "小组"]}` |
| `horizontal_match` | 横向合并 | `{"type": "horizontal_match", "targets": ["Q1", "Q2", "Q3", "Q4"]}` |

**纵向合并父子级联**：多个列会按顺序形成层级关系，父列变化会强制子列重新开始合并。

---

### 4. 样式库定义

样式库支持完整的 Excel 样式属性：

```json
{
  "header": {
    "font_name": "Microsoft YaHei",
    "font_size": 11,
    "bold": true,
    "bg_color": "#BFBFBF",
    "border": "thin",
    "align": "center",
    "valign": "vcenter"
  },
  "money": {
    "num_format": "#,##0.00",
    "font_color": "#006100"
  },
  "warning": {
    "bg_color": "#FFC7CE",
    "font_color": "#9C0006",
    "bold": true
  }
}
```

**支持的样式属性**：

- **字体**：`font_name`, `font_size`, `font_color` (#RRGGBB), `bold`, `italic`, `underline`
- **对齐**：`align` (left/center/right/fill/justify), `valign` (top/vcenter/bottom), `text_wrap`
- **边框**：`border` (none/thin/medium/thick/dashed/double), `border_color`
- **背景**：`bg_color`, `fg_color`, `pattern` (0-18)
- **数字格式**：`num_format` (如 `"yyyy-mm-dd"`, `"#,##0.00"`)

---

### 5. 读取 Excel 文件

支持读取 .xlsx 和 .xls 格式：

```rust
use xlsx_writer::prelude::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 读取指定工作表
    let wb = Workbook::new()?
        .read("input.xlsx", vec![
            ReadSheet::new("Sheet1".into()),
        ]);

    // 获取工作表并转换为 DataFrame
    if let Some(sheet) = wb.sheet("Sheet1") {
        let df = sheet.to_dataframe()?;
        println!("{:?}", df);
    }

    Ok(())
}
```

**读取配置选项**：

```rust
ReadSheet {
    sheet_name: "Sheet1".into(),
    skip_rows: Some(2),                    // 跳过前2行
    force_string_cols: Some(vec!["编号".into()]),  // 强制指定列为字符串
}
```

---

### 6. 双格式保存

根据文件扩展名自动选择格式：

```rust
// 保存为 .xlsx（推荐，保留完整样式）
workbook.save("output.xlsx")?;

// 保存为 .xls（兼容旧版 Excel，不保留样式）
workbook.save("output.xls")?;
```

**注意**：.xls 格式仅保留数据，样式信息将被忽略。

---

## 📋 完整配置示例

以下是一个完整的 JSON 配置示例，展示了样式规则、合并规则和样式库的组合使用：

```json
{
  "style_rules": [
    {
      "row_conditions": [{"type": "index", "criteria": [0]}],
      "apply": {"style": "header"}
    },
    {
      "row_conditions": [
        {"type": "value_range", "targets": ["完成率"], "criteria": ">=100"}
      ],
      "apply": {
        "style": "success",
        "overrides": [
          {"style": "percent_format", "col_conditions": [{"type": "match", "targets": ["完成率"], "criteria": ["完成率"]}]}
        ]
      }
    }
  ],
  "merge_rules": [
    {"type": "vertical_match", "targets": ["部门"]},
    {"type": "fixed", "targets": [[0, 0, 0, 3]]}
  ]
}
```

对应的样式库：

```json
{
  "header": {
    "font_name": "Microsoft YaHei",
    "font_size": 11,
    "bold": true,
    "bg_color": "#BFBFBF",
    "border": "thin",
    "align": "center"
  },
  "success": {
    "bg_color": "#C6EFCE",
    "font_color": "#006100"
  },
  "percent_format": {
    "num_format": "0.00%"
  }
}
```

---

## 🎯 典型使用场景

### 场景 1：自动化报表生成

数据分析师使用 Python/Rust 进行数据处理，通过 JSON 配置文件控制输出格式，无需修改代码即可调整报表样式。

### 场景 2：定时数据导出

在自动化任务中，使用固定的 JSON 配置文件，每次运行自动应用一致的格式规范。

### 场景 3：遗留系统兼容

需要同时支持现代 .xlsx 和旧版 .xls 格式的企业环境。

---

## ⚠️ 注意事项

1. **Sheet 名称限制**：
   - 最大 31 个字符
   - 不能包含 `\ / ? * : [ ]` 字符
   - 非法名称会自动回退为 "Sheet N" 格式

2. **.xls 格式限制**：
   - 仅保留数据，不保留样式
   - 行数限制为 65536 行
   - 列数限制为 256 列

3. **依赖版本**：
   - Polars 版本需与库保持一致（0.53.0）
   - 建议在 `Cargo.lock` 中锁定版本

4. **内存使用**：
   - 大型 DataFrame 建议分片处理
   - 样式映射使用 `Arc<str>` 共享内存，减少重复分配

---

## 🔗 相关链接

- [Polars 文档](https://docs.rs/polars/)
- [rust_xlsxwriter 文档](https://docs.rs/rust_xlsxwriter/)
- [BIFF8 格式规范](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/)

---

## 📄 License

MIT License
