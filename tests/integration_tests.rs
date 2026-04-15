//! 集成测试
//!
//! 测试场景：
//! 1. 销售数据 + 样式工厂
//! 2. 部门层级 + 合并工厂
//! 3. 多区域复杂报表
//! 4. 完整配置驱动

use polars::prelude::*;
use serde_json::json;
use std::fs;
use std::path::Path;
use std::sync::Arc;
use xlsx_writer::prelude::*;
use xlsx_writer::{Cell, RegionFactory, RegionStyles, SheetRegion};

const DATA_DIR: &str = "data";
const ASSETS_DIR: &str = "test_assets";

/// 设置测试环境
fn setup() {
    fs::create_dir_all(DATA_DIR).expect("Failed to create data directory");
}

/// 清理旧文件（保留最近5个）
fn cleanup_old_files(prefix: &str) {
    let mut files: Vec<_> = fs::read_dir(DATA_DIR)
        .unwrap_or_else(|_| panic!("Failed to read data directory"))
        .filter_map(|entry| entry.ok())
        .filter(|entry| {
            entry
                .file_name()
                .to_str()
                .map(|name| name.starts_with(prefix))
                .unwrap_or(false)
        })
        .collect();

    // 按修改时间排序（最新的在前）
    files.sort_by(|a, b| {
        b.metadata()
            .and_then(|m| m.modified())
            .unwrap_or(std::time::UNIX_EPOCH)
            .cmp(
                &a.metadata()
                    .and_then(|m| m.modified())
                    .unwrap_or(std::time::UNIX_EPOCH),
            )
    });

    // 删除超过5个的旧文件
    for entry in files.iter().skip(5) {
        let _ = fs::remove_file(entry.path());
    }
}

/// 测试1：销售数据 + RegionFactory（推荐使用）
#[test]
fn test_sales_with_styles() {
    setup();

    // 创建销售数据
    let df = df! {
        "部门" => &["技术部", "技术部", "技术部", "销售部", "销售部", "财务部", "财务部", "运营部", "运营部", "人事部"],
        "销售员" => &["张三", "李四", "王五", "赵六", "钱七", "孙八", "周九", "吴十", "郑十一", "王十二"],
        "Q1销售额" => &[120000.0, 80000.0, 150000.0, 200000.0, 95000.0, 60000.0, 110000.0, 130000.0, 75000.0, 85000.0],
        "Q2销售额" => &[135000.0, 85000.0, 165000.0, 220000.0, 105000.0, 65000.0, 120000.0, 145000.0, 80000.0, 90000.0],
        "完成率" => &[1.2, 0.95, 1.35, 1.5, 0.98, 0.87, 1.1, 1.25, 0.92, 1.05]
    }
    .expect("Failed to create DataFrame");

    println!("原始数据:");
    println!("{:?}", df);

    // 读取 JSON 配置
    let config_str = fs::read_to_string(format!("{}/sales_styles.json", ASSETS_DIR))
        .expect("Failed to read sales_styles.json");
    let config: serde_json::Value =
        serde_json::from_str(&config_str).expect("Failed to parse sales_styles.json");

    // 使用 RegionFactory 统一生成样式（推荐方式）
    let factory = RegionFactory::from_json(&config).expect("Failed to create RegionFactory");
    let styles = factory.execute(&df).expect("Failed to execute factory");

    // 创建 Region
    let region = SheetRegion::from_dataframe(df, "sales_data", Some(true), styles)
        .expect("Failed to create SheetRegion");

    println!("\nRegion 可视化:");
    println!("{}", region.visualize());

    // 创建工作簿并保存
    let output_path = format!("{}/test_sales_styles.xlsx", DATA_DIR);
    Workbook::new()
        .expect("Failed to create workbook")
        .with_library_from_json(&config["styles"])
        .expect("Failed to load style library")
        .add_sheet(WorkSheet::new("销售报表", vec![region]).expect("Failed to create worksheet"))
        .save(&output_path)
        .expect("Failed to save workbook");

    // 验证文件存在
    assert!(Path::new(&output_path).exists(), "Output file should exist");
    let metadata = fs::metadata(&output_path).expect("Failed to read file metadata");
    assert!(metadata.len() > 0, "Output file should not be empty");

    println!("\n✅ 测试1通过: 销售数据样式测试 - {}", output_path);

    cleanup_old_files("test_sales_styles");
}

/// 测试2：部门层级 + 合并工厂（生成 .xlsx 和 .xls）
#[test]
fn test_dept_merge() {
    setup();

    // 创建部门层级数据
    let df = df! {
        "部门" => &["技术部", "技术部", "技术部", "技术部", "技术部", "技术部",
                    "销售部", "销售部", "销售部", "销售部",
                    "财务部", "财务部"],
        "小组" => &["前端组", "前端组", "后端组", "后端组", "测试组", "测试组",
                    "华东区", "华东区", "华北区", "华北区",
                    "会计组", "出纳组"],
        "成员" => &["张三", "李四", "王五", "赵六", "孙七", "周八",
                    "吴九", "郑十", "钱十一", "陈十二",
                    "刘十三", "黄十四"],
        "业绩" => &[100, 120, 150, 130, 90, 110, 200, 180, 220, 190, 80, 85]
    }
    .expect("Failed to create DataFrame");

    println!("原始数据:");
    println!("{:?}", df);

    // 读取 JSON 配置
    let config_str = fs::read_to_string(format!("{}/dept_merge.json", ASSETS_DIR))
        .expect("Failed to read dept_merge.json");
    let config: serde_json::Value =
        serde_json::from_str(&config_str).expect("Failed to parse dept_merge.json");

    // 使用 RegionFactory 统一生成样式和合并区域（推荐方式）
    let factory = RegionFactory::from_json(&config).expect("Failed to create RegionFactory");
    let styles = factory.execute(&df).expect("Failed to execute factory");

    println!("生成的合并区域: {:?}", styles.merge_ranges);

    // 创建 Region
    let region = SheetRegion::from_dataframe(df, "dept_data", Some(true), styles)
        .expect("Failed to create SheetRegion");

    println!("\nRegion 可视化:");
    println!("{}", region.visualize_compact());

    // 测试 .xlsx 格式
    let output_xlsx = format!("{}/test_dept_merge.xlsx", DATA_DIR);
    Workbook::new()
        .expect("Failed to create workbook")
        .with_library_from_json(&config["styles"])
        .expect("Failed to load style library")
        .add_sheet(
            WorkSheet::new("部门层级", vec![region.clone()]).expect("Failed to create worksheet"),
        )
        .save(&output_xlsx)
        .expect("Failed to save xlsx");

    assert!(Path::new(&output_xlsx).exists(), "XLSX file should exist");
    println!("\n✅ 测试2-1通过: 部门合并 .xlsx - {}", output_xlsx);

    // 测试 .xls 格式
    let output_xls = format!("{}/test_dept_merge.xls", DATA_DIR);
    Workbook::new()
        .expect("Failed to create workbook")
        .add_sheet(WorkSheet::new("部门层级", vec![region]).expect("Failed to create worksheet"))
        .save(&output_xls)
        .expect("Failed to save xls");

    assert!(Path::new(&output_xls).exists(), "XLS file should exist");
    println!("✅ 测试2-2通过: 部门合并 .xls - {}", output_xls);

    cleanup_old_files("test_dept_merge");
}

/// 测试3：多区域复杂报表
#[test]
fn test_multi_region() {
    setup();

    // Region 1: 报表标题
    let header_data = vec![vec![Some(Cell::Text("2024年第二季度销售报表".to_string()))]];
    let mut title_styles = RegionStyles::new();
    title_styles
        .cell_styles
        .insert((0, 0), std::sync::Arc::from("title"));
    let header_region = SheetRegion::new("header", header_data)
        .with_styles(title_styles)
        .with_merge_ranges(vec![(0, 0, 0, 3)]);

    // Region 2: 汇总数据
    let summary_df = df! {
        "指标" => &["总销售额", "总订单数", "平均客单价", "完成率"],
        "数值" => &["1,250,000", "3,450", "362.32", "108.5%"]
    }
    .expect("Failed to create summary DataFrame");
    let summary_region =
        SheetRegion::from_dataframe(summary_df, "summary", Some(true), RegionStyles::new())
            .expect("Failed to create summary region");

    // Region 3: 详细数据（带合并和复杂样式）
    let detail_df = df! {
        "日期" => &["2024-04-01", "2024-04-02", "2024-04-03", "2024-04-04", "2024-04-05"],
        "产品" => &["产品A", "产品B", "产品A", "产品C", "产品B"],
        "数量" => &[100, 150, 80, 200, 120],
        "金额" => &[50000.0, 75000.0, 40000.0, 100000.0, 60000.0]
    }
    .expect("Failed to create detail DataFrame");

    let mut detail_styles = RegionStyles::new();
    detail_styles
        .cell_styles
        .insert((2, 3), std::sync::Arc::from("highlight"));
    detail_styles
        .cell_styles
        .insert((4, 3), std::sync::Arc::from("highlight"));

    let detail_region = SheetRegion::from_dataframe(detail_df, "detail", Some(true), detail_styles)
        .expect("Failed to create detail region");

    // 创建样式库
    let styles = serde_json::json!({
        "title": {
            "font_name": "Microsoft YaHei",
            "font_size": 16,
            "bold": true,
            "bg_color": "#4472C4",
            "font_color": "#FFFFFF",
            "align": "center"
        },
        "summary_header": {
            "font_name": "Microsoft YaHei",
            "font_size": 11,
            "bold": true,
            "bg_color": "#5B9BD5",
            "font_color": "#FFFFFF"
        },
        "detail_header": {
            "font_name": "Microsoft YaHei",
            "font_size": 11,
            "bold": true,
            "bg_color": "#B4C7DC"
        },
        "highlight": {
            "bg_color": "#FFE699",
            "bold": true
        },
        "money": {
            "num_format": "#,##0.00",
            "align": "right"
        }
    });

    // 创建多区域工作表
    let worksheet = WorkSheet::new(
        "多区域报表",
        vec![header_region, summary_region, detail_region],
    )
    .expect("Failed to create worksheet");

    println!("工作表包含 {} 个区域", worksheet.regions.len());
    for (i, region) in worksheet.regions.iter().enumerate() {
        println!("\n区域 {}: {}", i + 1, region.name);
        println!("{}", region.visualize_compact());
    }

    // 保存
    let output_path = format!("{}/test_multi_region.xlsx", DATA_DIR);
    Workbook::new()
        .expect("Failed to create workbook")
        .with_library_from_json(&styles)
        .expect("Failed to load style library")
        .add_sheet(worksheet)
        .save(&output_path)
        .expect("Failed to save workbook");

    assert!(Path::new(&output_path).exists(), "Output file should exist");
    println!("\n✅ 测试3通过: 多区域报表 - {}", output_path);

    cleanup_old_files("test_multi_region");
}

/// 测试4：完整配置驱动（零代码修改）
#[test]
fn test_full_config() {
    setup();

    // 创建测试数据
    let df = df! {
        "项目名称" => &["项目A", "项目B", "项目C", "项目D", "项目E", "项目F", "项目G", "项目H"],
        "负责人" => &["张三", "李四", "王五", "赵六", "钱七", "孙八", "周九", "吴十"],
        "预算" => &[50000.0, 80000.0, 120000.0, 45000.0, 95000.0, 60000.0, 150000.0, 70000.0],
        "实际支出" => &[48000.0, 85000.0, 115000.0, 42000.0, 98000.0, 58000.0, 155000.0, 72000.0],
        "状态" => &["正常", "超支", "正常", "正常", "超支", "正常", "超支", "正常"]
    }
    .expect("Failed to create DataFrame");

    println!("原始数据:");
    println!("{:?}", df);

    // 读取完整 JSON 配置
    let config_str = fs::read_to_string(format!("{}/complex_config.json", ASSETS_DIR))
        .expect("Failed to read complex_config.json");
    let config: serde_json::Value =
        serde_json::from_str(&config_str).expect("Failed to parse complex_config.json");

    // 创建 StyleFactory 和 MergeFactory
    let style_factory =
        StyleFactory::new(config["style_rules"].clone()).expect("Failed to create StyleFactory");
    let merge_factory =
        MergeFactory::new(config["merge_rules"].clone()).expect("Failed to create MergeFactory");

    // 执行规则
    let cell_styles = style_factory
        .execute(&df)
        .expect("Failed to execute styles");
    let merge_ranges = merge_factory
        .execute(&df)
        .expect("Failed to execute merges");

    println!("生成的样式条目数: {}", cell_styles.len());
    println!("生成的合并区域: {:?}", merge_ranges);

    // 创建 RegionStyles
    let mut styles = RegionStyles::new();
    styles.cell_styles = cell_styles;
    styles.merge_ranges = merge_ranges;

    // 创建 Region
    let region = SheetRegion::from_dataframe(df, "project_data", Some(true), styles)
        .expect("Failed to create SheetRegion");

    println!("\nRegion 验证报告:");
    let issues = region.validate();
    if issues.is_empty() {
        println!("✅ 无问题");
    } else {
        for issue in &issues {
            println!("⚠️  {}", issue);
        }
    }

    println!("\nRegion 可视化:");
    println!("{}", region.visualize());

    // 创建工作簿
    let output_path = format!("{}/test_full_config.xlsx", DATA_DIR);
    Workbook::new()
        .expect("Failed to create workbook")
        .with_library_from_json(&config["styles"])
        .expect("Failed to load style library")
        .add_sheet(WorkSheet::new("项目报表", vec![region]).expect("Failed to create worksheet"))
        .save(&output_path)
        .expect("Failed to save workbook");

    assert!(Path::new(&output_path).exists(), "Output file should exist");
    let metadata = fs::metadata(&output_path).expect("Failed to read file metadata");
    assert!(metadata.len() > 0, "Output file should not be empty");

    println!("\n✅ 测试4通过: 完整配置驱动 - {}", output_path);
    println!("文件大小: {} bytes", metadata.len());

    cleanup_old_files("test_full_config");
}

/// 测试5：无表头模式
#[test]
fn test_no_header() {
    setup();

    let df = df! {
        "列1" => &[1, 2, 3, 4, 5],
        "列2" => &["A", "B", "C", "D", "E"],
        "列3" => &[100.0, 200.0, 300.0, 400.0, 500.0]
    }
    .expect("Failed to create DataFrame");

    // 创建显式样式映射，给第1行数据（原坐标，包含表头时）设置自定义样式
    // 当 include_header=false 时，这个样式会调整到第0行
    let mut styles = RegionStyles::new();
    styles.cell_styles.insert((1, 0), Arc::from("custom_data")); // 第1行数据（原坐标）

    // 测试 include_header = false
    let region = SheetRegion::from_dataframe(df, "no_header_data", Some(false), styles)
        .expect("Failed to create SheetRegion");

    println!("无表头模式 Region:");
    println!("行数: {}", region.row_count());
    println!("列数: {}", region.col_count());
    println!("列名: {:?}", region.column_names());
    println!("{}", region.visualize_compact());

    // 验证：无表头时，第0行应该是数据而不是列名
    assert_eq!(region.row_count(), 5, "Should have 5 data rows");

    // 验证：检查样式坐标调整是否正确
    // include_header=false 时，原来的 (1,0) 样式应该调整到 (0,0)
    assert!(
        region.get_style(0, 0).is_some(),
        "Style should be at (0,0) for first data row"
    );
    assert_eq!(region.get_style(0, 0).unwrap().as_ref(), "custom_data");
    assert_eq!(region.get_style(0, 0).unwrap().as_ref(), "custom_data");

    // 创建工作簿时添加样式库
    let styles = json!({
        "custom_data": {
            "font_color": "#FF0000", // 红色字体
            "bold": true
        }
    });

    let output_path = format!("{}/test_no_header.xlsx", DATA_DIR);
    Workbook::new()
        .expect("Failed to create workbook")
        .with_library_from_json(&styles)
        .expect("Failed to load style library")
        .add_sheet(WorkSheet::new("无表头测试", vec![region]).expect("Failed to create worksheet"))
        .save(&output_path)
        .expect("Failed to save workbook");

    assert!(Path::new(&output_path).exists(), "Output file should exist");
    println!("\n✅ 测试5通过: 无表头模式（含样式验证）- {}", output_path);

    cleanup_old_files("test_no_header");
}

/// 测试6：最复杂场景 - 多 Sheet、多 Region、样式和合并
#[test]
fn test_complex_multi_sheet() {
    setup();

    // ========== Sheet 1: 复杂多 Region 报表 ==========

    // Region 1: 标题（合并单元格）
    let title_data = vec![vec![Some(Cell::Text("2024年度综合报表".to_string()))]];
    let mut title_styles = RegionStyles::new();
    title_styles.cell_styles.insert((0, 0), Arc::from("title"));
    let title_region = SheetRegion::new("title", title_data)
        .with_styles(title_styles)
        .with_merge_ranges(vec![(0, 0, 0, 5)]); // 合并6列

    // Region 2: 汇总统计（带条件样式）
    let summary_df = df! {
        "指标" => &["总销售额", "总订单", "平均客单价", "完成率", "增长率", "客户满意度"],
        "数值" => &[2500000.0, 8500.0, 294.12, 1.15, 0.23, 0.92]
    }
    .expect("Failed to create summary DataFrame");

    let summary_region =
        SheetRegion::from_dataframe(summary_df, "summary", Some(true), RegionStyles::new())
            .expect("Failed to create summary region");

    // Region 3: 详细数据（带合并和复杂样式）
    let detail_df = df! {
        "区域" => &["华东", "华东", "华东", "华北", "华北", "华南", "华南", "华南"],
        "城市" => &["上海", "南京", "杭州", "北京", "天津", "广州", "深圳", "佛山"],
        "Q1" => &[100000, 80000, 90000, 120000, 70000, 110000, 130000, 60000],
        "Q2" => &[120000, 85000, 95000, 135000, 75000, 125000, 145000, 65000],
        "Q3" => &[110000, 82000, 92000, 130000, 72000, 120000, 140000, 62000],
        "Q4" => &[130000, 88000, 98000, 140000, 78000, 135000, 155000, 68000]
    }
    .expect("Failed to create detail DataFrame");

    let mut detail_styles = RegionStyles::new();
    // 高亮高销售额
    detail_styles
        .cell_styles
        .insert((1, 2), Arc::from("high_sales")); // 上海 Q1
    detail_styles
        .cell_styles
        .insert((4, 2), Arc::from("high_sales")); // 北京 Q1
    detail_styles
        .cell_styles
        .insert((6, 2), Arc::from("high_sales")); // 深圳 Q1

    // 将 merge_ranges 放入 detail_styles
    detail_styles.merge_ranges = vec![
        (1, 0, 3, 0), // 华东合并
        (4, 0, 5, 0), // 华北合并
        (6, 0, 8, 0), // 华南合并
    ];

    let detail_region = SheetRegion::from_dataframe(detail_df, "detail", Some(true), detail_styles)
        .expect("Failed to create detail region");

    let sheet1 = WorkSheet::new(
        "综合报表",
        vec![title_region, summary_region, detail_region],
    )
    .expect("Failed to create sheet1");

    // ========== Sheet 2: 另一个复杂报表 ==========

    // Region 1: 部门汇总
    let dept_df = df! {
        "部门" => &["技术部", "技术部", "销售部", "销售部", "财务部"],
        "小组" => &["前端", "后端", "华东组", "华北组", "会计"],
        "人数" => &[15, 20, 25, 20, 8],
        "预算" => &[500000, 600000, 800000, 700000, 200000],
        "支出" => &[480000, 580000, 820000, 680000, 190000]
    }
    .expect("Failed to create dept DataFrame");

    // 创建包含 merge_ranges 的 RegionStyles
    let mut dept_styles = RegionStyles::new();
    dept_styles.merge_ranges = vec![
        (1, 0, 2, 0), // 技术部合并
        (3, 0, 4, 0), // 销售部合并
    ];

    let dept_region = SheetRegion::from_dataframe(dept_df, "dept", Some(true), dept_styles)
        .expect("Failed to create dept region");

    // Region 2: 备注说明（合并单元格）
    let note_data = vec![vec![Some(Cell::Text(
        "注：以上数据截至2024年12月31日，未经审计。".to_string(),
    ))]];
    let mut note_styles = RegionStyles::new();
    note_styles.cell_styles.insert((0, 0), Arc::from("note"));
    let note_region = SheetRegion::new("note", note_data)
        .with_styles(note_styles)
        .with_merge_ranges(vec![(0, 0, 0, 4)]);

    let sheet2 = WorkSheet::new("部门统计", vec![dept_region, note_region])
        .expect("Failed to create sheet2");

    // ========== 创建完整样式库 ==========
    let styles = json!({
        "title": {
            "font_name": "Microsoft YaHei",
            "font_size": 16,
            "bold": true,
            "bg_color": "#4472C4",
            "font_color": "#FFFFFF",
            "align": "center"
        },
        "high_sales": {
            "bg_color": "#FFC7CE",
            "font_color": "#9C0006",
            "bold": true
        },
        "note": {
            "font_size": 9,
            "italic": true,
            "font_color": "#666666"
        }
    });

    // ========== 保存工作簿 ==========
    let output_path = format!("{}/test_complex_multi_sheet.xlsx", DATA_DIR);
    Workbook::new()
        .expect("Failed to create workbook")
        .with_library_from_json(&styles)
        .expect("Failed to load style library")
        .add_sheet(sheet1)
        .add_sheet(sheet2)
        .save(&output_path)
        .expect("Failed to save workbook");

    assert!(Path::new(&output_path).exists(), "Output file should exist");
    let metadata = fs::metadata(&output_path).expect("Failed to read file metadata");
    assert!(metadata.len() > 0, "Output file should not be empty");

    println!("\n✅ 复杂多 Sheet 测试通过: {}", output_path);
    println!("文件大小: {} bytes", metadata.len());
    println!("包含 2 个 Sheet，共 5 个 Region，多处样式和合并");

    cleanup_old_files("test_complex_multi_sheet");
}

#[cfg(test)]
mod dimension_factory_tests {
    use super::*;
    use xlsx_writer::dimension_factory::DimensionFactory;

    /// 测试基本的行高设置
    #[test]
    fn test_dimension_factory_fixed_row_height() {
        let df = df! {
            "name" => ["Alice", "Bob", "Charlie"],
            "age" => [25, 30, 35]
        }
        .expect("Failed to create DataFrame");

        let config = json!([
            {
                "target": "row",
                "condition": {"type": "index", "criteria": [0, 2]},
                "value": {"type": "fixed", "value": 30.0}
            }
        ]);

        let factory = DimensionFactory::new(config).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        assert_eq!(result.row_heights.len(), 2);
        assert_eq!(result.row_heights.get(&0), Some(&30.0));
        assert_eq!(result.row_heights.get(&2), Some(&30.0));
        assert!(result.row_heights.get(&1).is_none());
    }

    /// 测试基本的列宽设置
    #[test]
    fn test_dimension_factory_fixed_col_width() {
        let df = df! {
            "name" => ["Alice", "Bob"],
            "age" => [25, 30]
        }
        .expect("Failed to create DataFrame");

        let config = json!([
            {
                "target": "column",
                "condition": {"type": "index", "criteria": [0]},
                "value": {"type": "fixed", "value": 20.0}
            }
        ]);

        let factory = DimensionFactory::new(config).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        assert_eq!(result.col_widths.len(), 1);
        assert_eq!(result.col_widths.get(&0), Some(&20.0));
        assert!(result.col_widths.get(&1).is_none());
    }

    /// 测试列宽 Auto 计算（中文宽度）
    #[test]
    fn test_dimension_factory_auto_col_width_chinese() {
        // 使用较长的字符串以确保超过最小宽度（8.0）
        // 中文：4个字符 * 2宽度 + 2padding = 10
        // 英文：8个字符 * 1宽度 + 2padding = 10，但我们需要英文更短一些
        let df = df! {
            "name" => ["张三李四", "王五赵六"],  // 4个中文字符 = 8宽度
            "code" => ["ABC", "DEF"]             // 3个英文字符 = 3宽度
        }
        .expect("Failed to create DataFrame");

        let config = json!([
            {
                "target": "column",
                "condition": {"type": "index", "criteria": [0, 1]},
                "value": {"type": "auto"}
            }
        ]);

        let factory = DimensionFactory::new(config).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        assert_eq!(result.col_widths.len(), 2);
        // 中文字符应该比英文字符宽
        // 中文列：8 + 2 = 10（超过最小值8）
        // 英文列：3 + 2 = 5，被钳制到8
        let chinese_width = result.col_widths.get(&0).expect("Missing column 0");
        let english_width = result.col_widths.get(&1).expect("Missing column 1");
        assert!(
            chinese_width > english_width,
            "Chinese column width ({}) should be greater than English column width ({})",
            chinese_width,
            english_width
        );
    }

    /// 测试规则覆盖（后定义规则覆盖前面的规则）
    #[test]
    fn test_dimension_factory_rule_override() {
        let df = df! {
            "name" => ["Alice", "Bob"]
        }
        .expect("Failed to create DataFrame");

        let config = json!([
            {
                "target": "row",
                "condition": {"type": "index", "criteria": [0]},
                "value": {"type": "fixed", "value": 20.0}
            },
            {
                "target": "row",
                "condition": {"type": "index", "criteria": [0]},
                "value": {"type": "fixed", "value": 30.0}
            }
        ]);

        let factory = DimensionFactory::new(config).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        // 后面的规则应该覆盖前面的
        assert_eq!(result.row_heights.get(&0), Some(&30.0));
    }

    /// 测试负数索引
    #[test]
    fn test_dimension_factory_negative_index() {
        let df = df! {
            "name" => ["Alice", "Bob", "Charlie"]
        }
        .expect("Failed to create DataFrame");

        let config = json!([
            {
                "target": "row",
                "condition": {"type": "index", "criteria": [-1]},
                "value": {"type": "fixed", "value": 40.0}
            }
        ]);

        let factory = DimensionFactory::new(config).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        // -1 应该对应最后一行数据（Excel 物理行号 3）
        assert_eq!(result.row_heights.get(&3), Some(&40.0));
    }

    /// 测试 Match 条件定位列
    #[test]
    fn test_dimension_factory_match_condition_for_column() {
        let df = df! {
            "user_name" => ["Alice", "Bob"],
            "user_age" => [25, 30],
            "dept" => ["IT", "HR"]
        }
        .expect("Failed to create DataFrame");

        let config = json!([
            {
                "target": "column",
                "condition": {
                    "type": "match",
                    "targets": ["user_name", "user_age"],
                    "criteria": ["user_name", "user_age"]
                },
                "value": {"type": "fixed", "value": 25.0}
            }
        ]);

        let factory = DimensionFactory::new(config).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        // Match 条件应该匹配列名
        assert!(result.col_widths.get(&0).is_some()); // user_name
        assert!(result.col_widths.get(&1).is_some()); // user_age
        assert!(result.col_widths.get(&2).is_none()); // dept
    }

    /// 测试 ValueRange 条件在列宽规则中被静默忽略
    #[test]
    fn test_dimension_factory_value_range_ignored_for_column() {
        let df = df! {
            "age" => [25, 30, 35]
        }
        .expect("Failed to create DataFrame");

        let config = json!([
            {
                "target": "column",
                "condition": {
                    "type": "value_range",
                    "targets": ["age"],
                    "criteria": ">30"
                },
                "value": {"type": "fixed", "value": 20.0}
            }
        ]);

        let factory = DimensionFactory::new(config).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        // ValueRange 在列宽规则中应该被静默忽略
        assert!(result.col_widths.is_empty());
    }

    /// 测试空规则集
    #[test]
    fn test_dimension_factory_empty_rules() {
        let df = df! {
            "name" => ["Alice", "Bob"]
        }
        .expect("Failed to create DataFrame");

        let factory = DimensionFactory::new(json!([])).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        assert!(result.row_heights.is_empty());
        assert!(result.col_widths.is_empty());
    }

    /// 测试 JSON 字符串创建工厂
    #[test]
    fn test_dimension_factory_from_json_str() {
        let json_str = r#"[{"target": "row", "condition": {"type": "index", "criteria": [0]}, "value": {"type": "fixed", "value": 35.0}}]"#;

        let factory = DimensionFactory::from_json_str(json_str).expect("Failed to create factory");

        let df = df! {
            "name" => ["Alice", "Bob"]
        }
        .expect("Failed to create DataFrame");

        let result = factory.execute(&df).expect("Failed to execute");

        assert_eq!(result.row_heights.get(&0), Some(&35.0));
    }

    /// 测试 All 条件 - 为所有列启用自动列宽
    #[test]
    fn test_dimension_factory_all_columns_auto_width() {
        let df = df! {
            "name" => ["Alice", "Bob", "Charlie"],
            "age" => [25, 30, 35],
            "city" => ["New York", "Los Angeles", "Chicago"]
        }
        .expect("Failed to create DataFrame");

        let config = json!([
            {
                "target": "column",
                "condition": {"type": "all"},
                "value": {"type": "auto"}
            }
        ]);

        let factory = DimensionFactory::new(config).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        // 应该为所有3列都计算了宽度
        assert_eq!(result.col_widths.len(), 3, "All columns should have widths");
        assert!(
            result.col_widths.contains_key(&0),
            "Column 0 should have width"
        );
        assert!(
            result.col_widths.contains_key(&1),
            "Column 1 should have width"
        );
        assert!(
            result.col_widths.contains_key(&2),
            "Column 2 should have width"
        );
    }

    /// 测试 All 条件 - 为所有行设置固定行高
    #[test]
    fn test_dimension_factory_all_rows_fixed_height() {
        let df = df! {
            "name" => ["Alice", "Bob", "Charlie"]
        }
        .expect("Failed to create DataFrame");

        let config = json!([
            {
                "target": "row",
                "condition": {"type": "all"},
                "value": {"type": "fixed", "value": 25.0}
            }
        ]);

        let factory = DimensionFactory::new(config).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        // 应该为所有4行都设置了行高（包括 header 行 0 和 3行数据 1,2,3）
        assert_eq!(result.row_heights.len(), 4, "All rows should have heights");
        assert_eq!(result.row_heights.get(&0), Some(&25.0)); // header 行
        assert_eq!(result.row_heights.get(&1), Some(&25.0)); // 数据行 0
        assert_eq!(result.row_heights.get(&2), Some(&25.0)); // 数据行 1
        assert_eq!(result.row_heights.get(&3), Some(&25.0)); // 数据行 2
    }

    /// 测试 All 条件与其他条件的覆盖关系
    #[test]
    fn test_dimension_factory_all_with_override() {
        let df = df! {
            "name" => ["Alice", "Bob", "Charlie"]
        }
        .expect("Failed to create DataFrame");

        let config = json!([
            {
                "target": "row",
                "condition": {"type": "all"},
                "value": {"type": "fixed", "value": 20.0}
            },
            {
                "target": "row",
                "condition": {"type": "index", "criteria": [1]},
                "value": {"type": "fixed", "value": 35.0}
            }
        ]);

        let factory = DimensionFactory::new(config).expect("Failed to create factory");
        let result = factory.execute(&df).expect("Failed to execute");

        // 所有行都应该有行高（4行：header 0 + 数据 1,2,3）
        assert_eq!(result.row_heights.len(), 4);
        // 第0行(header)和第2,3行使用 All 设置的值
        assert_eq!(result.row_heights.get(&0), Some(&20.0)); // header 行
        assert_eq!(result.row_heights.get(&2), Some(&20.0)); // 数据行 1
        assert_eq!(result.row_heights.get(&3), Some(&20.0)); // 数据行 2
                                                             // 第1行被后面的规则覆盖
        assert_eq!(result.row_heights.get(&1), Some(&35.0)); // 数据行 0
    }
}
