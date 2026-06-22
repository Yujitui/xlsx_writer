//! 样式工厂条件评估的回归测试

use polars::prelude::*;
use xlsx_writer::style_factory::{evaluate_row_conditions, StyleCondition};

/// 验证 Equal 条件在 criteria=false 时不会错误命中表头行。
///
/// 这是针对以下 BUG 的回归测试：
/// 当使用 `"type": "equal", "criteria": false` 时，代码会对 mask 取反，
/// 导致原本用于排除表头行的 mask[0] = false 被翻转为 true，
/// 从而使表头行被错误地判定为命中「不全相等」条件。
#[test]
fn test_equal_not_equal_condition_excludes_header() {
    let df = df! {
        "缴纳单位" => &["缴纳单位", "A公司", "B公司"],
        "成本单位" => &["成本单位", "A公司", "C公司"]
    }
    .expect("Failed to create DataFrame");

    let condition = StyleCondition::Equal {
        targets: vec!["缴纳单位".to_string(), "成本单位".to_string()],
        criteria: false,
    };

    let mask = evaluate_row_conditions(&df, &[condition])
        .expect("Failed to evaluate conditions");

    let values: Vec<bool> = mask.into_iter().map(|opt| opt.unwrap_or(false)).collect();

    // mask 长度 = height + 1（包含表头行）
    // 第 0 行：表头行，应始终为 false
    // 第 1 行："缴纳单位" != "成本单位"，命中 criteria=false → true
    // 第 2 行："A公司" == "A公司"，不命中 → false
    // 第 3 行："B公司" != "C公司"，命中 criteria=false → true
    assert_eq!(values, vec![false, true, false, true]);
}

/// 验证 Equal 条件在 criteria=true 时表头行仍然被正确排除。
#[test]
fn test_equal_equal_condition_excludes_header() {
    let df = df! {
        "预算" => &[100.0, 100.0, 200.0],
        "实际支出" => &[100.0, 120.0, 200.0]
    }
    .expect("Failed to create DataFrame");

    let condition = StyleCondition::Equal {
        targets: vec!["预算".to_string(), "实际支出".to_string()],
        criteria: true,
    };

    let mask = evaluate_row_conditions(&df, &[condition])
        .expect("Failed to evaluate conditions");

    let values: Vec<bool> = mask.into_iter().map(|opt| opt.unwrap_or(false)).collect();

    // 第 0 行：表头行，应始终为 false
    // 第 1 行：100.0 == 100.0，命中 criteria=true → true
    // 第 2 行：100.0 != 120.0，不命中 → false
    // 第 3 行：200.0 == 200.0，命中 criteria=true → true
    assert_eq!(values, vec![false, true, false, true]);
}
