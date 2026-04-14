# xlsx_writer

一个专为数据分析自动化设计的 Rust Excel 处理库，支持 **JSON 配置驱动**的样式和合并规则，同时兼容 **.xlsx** 和 **.xls** 双格式。

## ✨ 主要特性

- **📝 JSON 配置驱动**：通过 JSON 文件定义样式规则和合并区域，无需修改代码即可调整输出格式
- **📊 Polars 原生集成**：与 Polars DataFrame 无缝协作，支持大规模数据处理
- **🎨 智能样式系统**：基于数据内容的条件样式（数值范围、字符串匹配、列间比较等）
- **🔗 自动合并区域**：支持纵向、横向自动合并，以及静态坐标合并
- **💾 双格式支持**：同时支持现代 .xlsx 格式和旧版 .xls (Excel 97-2003) 格式
- **🖥️ 跨平台字体**：自动适配 Windows（微软雅黑）和 macOS（苹方）的中文字体
- **🧩 多区域支持**：单个工作表支持多个独立数据区域，灵活组合复杂报表

## 📖 文档

| 文档 | 说明 |
|------|------|
| [📚 完整配置指南](docs/configuration_guide.md) | 系统学习所有配置选项（1700+行详细文档，含样式库配置） |
| [⚡ 快速参考手册](docs/QUICK_REFERENCE.md) | 常用配置模式速查（368行速查手册） |
| [📁 示例配置](docs/examples/) | 6个完整配置示例 |

**新手推荐**: 先阅读 [快速开始](docs/configuration_guide.md#1-快速开始) 章节，然后查看 [基础示例](docs/examples/basic.json)。

### 核心配置概念

```json
{
  "styles": {                    // 1️⃣ 定义样式库
    "header": { "bg_color": "#4472C4", "bold": true }
  },
  "style_rules": [...],          // 2️⃣ 样式应用规则（引用 styles）
  "merge_rules": [...],          // 3️⃣ 合并单元格规则
  "dimension_rules": [...]       // 4️⃣ 行高列宽规则
}
```

**使用方法**：
```rust
let workbook = Workbook::new()?
    .with_library_from_json(&config["styles"])?  // 加载样式库
    .with_dataframe(df, &config)?;               // 应用所有规则
```

## 📚 核心概念

### 工作表区域（SheetRegion）

**SheetRegion** 是数据导出的最小单元，包含：
- **数据**：二维单元格数组
- **样式映射**：单元格坐标 → 样式名称
- **合并区域**：需要合并的单元格范围

**Region 本身不预设数据结构**，第0行可以是表头、数据或任何内容。

```rust
use xlsx_writer::{Cell, SheetRegion};

// 手动创建：完全自由的结构
let region = SheetRegion::new("data", vec![
    vec![Some(Cell::Text("标题".into()))],  // 第0行
    vec![Some(Cell::Text("内容".into()))],  // 第1行
]);

// 从 DataFrame 创建：第0行自动使用列名作为表头
let region = SheetRegion::from_dataframe(
    df,
    "sales_data",
    Some(true),             // include_header: 是否将列名作为第0行
    Some(style_map),        // 样式映射
    Some(merge_ranges),     // 合并区域
)?;
```

### 坐标系统

**所有坐标都是 0-based**，从 Region 的左上角 (0, 0) 开始：

```
    col=0    col=1    col=2
row=0  [A]      [B]      [C]
row=1  [D]      [E]      [F]
row=2  [G]      [H]      [I]
```

### Factory 的坐标约定

**StyleFactory** 和 **MergeFactory** 基于 DataFrame 生成规则时，使用以下约定：

| Factory 坐标 | 对应 Region 行 | 说明 |
|-------------|---------------|------|
| row=0 | 第0行 | 预留为表头行 |
| row=1 | 第1行 | DataFrame 第0行数据 |
| row=2 | 第2行 | DataFrame 第1行数据 |
| ... | ... | ... |

**当 `include_header=false` 时**：
- Region 的第0行直接是数据（跳过列名）
- Factory 生成的坐标自动调整：row=1→row=0, row=2→row=1...
- row=0（原表头样式）被删除

### 多区域工作表

一个工作表可以包含多个 SheetRegion，按顺序排列写入 Excel：

```rust
let sheet = WorkSheet::new("报表", vec![
    title_region,    // Region 1
    data_region,     // Region 2  
    footer_region,   // Region 3
])?;
```

---

## 🚀 快速开始

### 安装

在 `Cargo.toml` 中添加依赖：

```toml
[dependencies]
xlsx_writer = { git = "https://github.com/yourusername/xlsx_writer" }
polars = { version = "0.53.0", features = ["lazy", "dtype-full"] }
serde_json = "1.0"
```

### 基础示例（多区域）

```rust
use xlsx_writer::prelude::*;
use xlsx_writer::{Cell, SheetRegion};
use polars::prelude::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 1. 准备数据
    let df = df!(
        "部门" => &["技术部", "技术部", "财务部", "财务部"],
        "姓名" => &["张三", "李四", "王五", "赵六"],
        "薪资" => &[15000.0, 18000.0, 12000.0, 14000.0]
    )?;

    // 2. 创建区域
    let region = SheetRegion::from_dataframe(
        df,
        "data",
        Some(true),     // 包含表头
        None,           // 无自定义样式
        None,           // 无合并
    )?;

    // 3. 创建并保存工作簿
    Workbook::new()?
        .add_sheet(WorkSheet::new("工资表", vec![region])?)
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

### 2. 区域样式管理

SheetRegion 提供完整的样式管理方法：

```rust
// 设置单个单元格样式
region.set_style(0, 0, "header");      // 第0行第0列
region.set_style(1, 2, "highlight");   // 第1行第2列

// 批量设置
region.set_row_style(0, "header");     // 整行
region.set_col_style(2, "money");      // 整列

// 清除样式
region.clear_style(0, 0);              // 清除单元格
region.clear_all_styles();             // 清除全部

// 查询样式
if let Some(style) = region.get_style(1, 2) {
    println!("样式: {}", style);
}

// 获取所有样式（调试）
println!("{}", region.visualize());
```

---

### 3. 区域合并管理

```rust
// 添加合并区域
region.add_merge(0, 0, 0, 3);          // 合并第0行的0-3列

// 批量添加（如从 Factory 生成）
region = region.with_merge_ranges(vec![
    (1, 0, 3, 0),  // 华东合并
    (4, 0, 5, 0),  // 华北合并
]);

// 查询合并
if region.is_merged(1, 0) {
    let merge = region.get_merge(1, 0);
    println!("合并区域: {:?}", merge);
}

// 清除合并
region.clear_merge_at(1, 0);           // 删除包含(1,0)的合并
region.clear_all_merges();             // 清除全部
```

---

### 5. 样式规则详解

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

### 6. 合并规则详解

自动计算并生成 Excel 合并区域：

#### 合并类型

| 类型 | 说明 | 示例 |
|------|------|------|
| `fixed` | 静态坐标 | `{"type": "fixed", "targets": [[0, 0, 0, 3]]}`（合并 A1:D1） |
| `vertical_match` | 纵向合并 | `{"type": "vertical_match", "targets": ["部门", "小组"]}` |
| `horizontal_match` | 横向合并 | `{"type": "horizontal_match", "targets": ["Q1", "Q2", "Q3", "Q4"]}` |

**纵向合并父子级联**：多个列会按顺序形成层级关系，父列变化会强制子列重新开始合并。

---

### 7. 样式库定义

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

### 8. 读取 Excel 文件

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

### 9. 复杂多区域报表

创建包含多个 Sheet、多个 Region 的复杂报表：

```rust
// Sheet 1: 综合报表
let title = SheetRegion::new("title", vec![vec![
    Some(Cell::Text("年度报表".into()))
]]).with_merge_ranges(vec![(0, 0, 0, 5)]);

let summary = SheetRegion::from_dataframe(summary_df, "summary", None, None, None)?;
let detail = SheetRegion::from_dataframe(detail_df, "detail", None, Some(styles), Some(merges))?;

let sheet1 = WorkSheet::new("综合报表", vec![title, summary, detail])?;

// Sheet 2: 部门统计
let dept = SheetRegion::from_dataframe(dept_df, "dept", None, None, Some(merges))?;
let note = SheetRegion::new("note", vec![vec![Some(Cell::Text("备注".into()))]]);

let sheet2 = WorkSheet::new("部门统计", vec![dept, note])?;

// 保存多 Sheet 工作簿
Workbook::new()?
    .with_library_from_json(&styles)?
    .add_sheet(sheet1)
    .add_sheet(sheet2)
    .save("complex_report.xlsx")?;
```

---

### 10. 双格式保存

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

1. **坐标系统**：
   - 所有样式和合并的坐标都是 **0-based**
   - Region 本身不预设第0行的属性（可以是表头、数据或任何内容）
   - 当 `include_header=false` 时，Factory 生成的坐标自动调整（row=1→row=0）

2. **Sheet 名称限制**：
   - 最大 31 个字符
   - 不能包含 `\ / ? * : [ ]` 字符
   - 非法名称会自动回退为 "Sheet N" 格式

3. **.xls 格式限制**：
   - 仅保留数据，不保留样式
   - 行数限制为 65536 行
   - 列数限制为 256 列

4. **Factory 生成规则**：
   - StyleFactory/MergeFactory 生成的坐标基于"表头+数据"坐标系统
   - row=0 对应表头，row=1 对应 DataFrame 第0行数据
   - 不需要表头时，row=0 的样式自动失效

5. **依赖版本**：
   - Polars 版本需与库保持一致（0.53.0）
   - 建议在 `Cargo.lock` 中锁定版本

6. **内存使用**：
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
