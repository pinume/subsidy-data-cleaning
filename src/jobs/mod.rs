use std::collections::HashMap;
use std::path::{Path, PathBuf};

use rust_decimal::Decimal;
use rust_decimal::prelude::FromPrimitive;

use crate::io::xlsx_reader::RawCell;
use crate::model::{ProcessError, Table, Value};
use crate::utils::{dates, doc_no, text};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Category {
    Invoice = 1,
    UnionSubsidy,
    UploadedDigital,
    UploadedAppliance,
    UnionPay,
    RefundAppliance,
    RefundDigital,
    Receipts,
    Coupons,
}

pub trait Job {
    fn category(&self) -> Category;
    fn title(&self) -> &'static str;
    fn output_stem(&self) -> &'static str;
    fn run(&self, input_dir: &Path) -> Result<Table, ProcessError>;
}

/// 文本字段：保留原值（包括纯空白），仅数值型单元格需还原为完整整数文本。
pub(crate) fn cell_text(cell: &RawCell) -> Result<String, String> {
    match cell {
        RawCell::Empty => Ok(String::new()),
        RawCell::Text(text) => Ok(text.clone()),
        RawCell::Int(n) => Ok(n.to_string()),
        RawCell::Bool(b) => Ok(b.to_string()),
        RawCell::Float(f) => text::identifier_from_float(*f)
            .ok_or_else(|| format!("数值 {f} 带小数或超出精度范围，无法还原为完整文本")),
        RawCell::DateTime(_) => Err("此字段不应为日期类型".to_string()),
        RawCell::Error(message) => Err(format!("单元格为错误值：{message}")),
    }
}

/// 空字符串转换为`Value::Empty`，否则包装为`Value::Text`。
pub(crate) fn text_value(value: String) -> Value {
    if value.is_empty() {
        Value::Empty
    } else {
        Value::Text(value)
    }
}

/// `None`转换为`Value::Empty`，否则包装为`Value::Decimal`。
pub(crate) fn amount_value(value: Option<Decimal>) -> Value {
    value.map_or(Value::Empty, Value::Decimal)
}

/// 数值字段：为空时保持为空，不做尾差取整（显示格式如`0.00`只影响展示，不改变实际值）。
/// 部分导出会把金额存成文本，按十进制文本解析（避免不必要的二进制浮点转换）。
pub(crate) fn cell_amount(cell: &RawCell) -> Result<Option<Decimal>, String> {
    match cell {
        RawCell::Empty => Ok(None),
        RawCell::Text(t) if t.trim().is_empty() => Ok(None),
        RawCell::Text(t) => t
            .trim()
            .parse::<Decimal>()
            .map(Some)
            .map_err(|_| format!("文本“{t}”无法解析为十进制金额")),
        RawCell::Int(n) => Ok(Some(Decimal::from(*n))),
        RawCell::Float(f) => Decimal::from_f64(*f)
            .map(Some)
            .ok_or_else(|| format!("数值 {f} 无法转换为十进制金额")),
        other => Err(format!("金额字段出现非数值内容：{other}")),
    }
}

/// 日期字段：源值为 Excel 日期序列值或`yyyymmdd`文本；无法识别时保留原值，不猜测修正
/// （第 5、6 节的通用豁免）。
pub(crate) fn cell_date_or_text(cell: &RawCell) -> Value {
    if let RawCell::DateTime(serial) = cell {
        return dates::date_from_serial(*serial)
            .map_or_else(|| Value::Text(cell.to_string()), Value::Date);
    }
    let text = cell.to_string();
    if text.trim().is_empty() {
        return Value::Empty;
    }
    dates::parse_yyyymmdd(&text).map_or(Value::Text(text), Value::Date)
}

/// 日期时间字段：保留到秒；无法识别时保留原值，不猜测修正（第 5、6 节的通用豁免）。
pub(crate) fn cell_datetime_or_text(cell: &RawCell) -> Value {
    if let RawCell::DateTime(serial) = cell {
        return dates::datetime_from_serial(*serial)
            .map_or_else(|| Value::Text(cell.to_string()), Value::DateTime);
    }
    let text = cell.to_string();
    if text.trim().is_empty() {
        return Value::Empty;
    }
    dates::parse_date_text(&text).map_or(Value::Text(text), Value::DateTime)
}

/// 数据异常错误的统一构造函数：文件名、工作表名、行号、字段名、原始值、异常说明。
pub(crate) fn data_error(
    file: &str,
    sheet: &str,
    row: u32,
    field: &str,
    value: String,
    detail: String,
) -> ProcessError {
    ProcessError::Data {
        file: file.to_string(),
        sheet: sheet.to_string(),
        row,
        field: field.to_string(),
        value,
        detail,
    }
}

/// 日期/日期时间/时间字段的共同骨架：为空时保持为空；有 Excel 序列值或文本但无法
/// 识别时按数据异常终止。`from_serial`/`from_text`/`to_value`承载各字段类型的差异。
#[allow(clippy::too_many_arguments)]
fn parse_temporal_field<T>(
    cell: &RawCell,
    field: &str,
    file: &str,
    sheet: &str,
    row: u32,
    err_label: &str,
    from_serial: impl FnOnce(f64) -> Option<T>,
    from_text: impl FnOnce(&str) -> Option<T>,
    to_value: impl FnOnce(T) -> Value,
) -> Result<Value, ProcessError> {
    match cell {
        RawCell::Empty => Ok(Value::Empty),
        RawCell::DateTime(serial) => from_serial(*serial).map(to_value).ok_or_else(|| {
            data_error(
                file,
                sheet,
                row,
                field,
                serial.to_string(),
                err_label.to_string(),
            )
        }),
        other => {
            let text = other.to_string();
            if text.trim().is_empty() {
                return Ok(Value::Empty);
            }
            from_text(&text)
                .map(to_value)
                .ok_or_else(|| data_error(file, sheet, row, field, text, err_label.to_string()))
        }
    }
}

/// 日期字段：为空时保持为空；非空但无法识别时按数据异常终止。
pub(crate) fn parse_date_field(
    cell: &RawCell,
    field: &str,
    file: &str,
    sheet: &str,
    row: u32,
) -> Result<Value, ProcessError> {
    parse_temporal_field(
        cell,
        field,
        file,
        sheet,
        row,
        "无法解析为日期",
        dates::date_from_serial,
        |text| dates::parse_date_text(text).map(|dt| dt.date()),
        Value::Date,
    )
}

/// 日期时间字段：为空时保持为空；非空但无法识别时按数据异常终止。
pub(crate) fn parse_datetime_field(
    cell: &RawCell,
    field: &str,
    file: &str,
    sheet: &str,
    row: u32,
) -> Result<Value, ProcessError> {
    parse_temporal_field(
        cell,
        field,
        file,
        sheet,
        row,
        "无法解析为日期时间",
        dates::datetime_from_serial,
        dates::parse_date_text,
        Value::DateTime,
    )
}

/// 时间字段：为空时保持为空；非空但无法识别时按数据异常终止。
pub(crate) fn parse_time_field(
    cell: &RawCell,
    field: &str,
    file: &str,
    sheet: &str,
    row: u32,
) -> Result<Value, ProcessError> {
    parse_temporal_field(
        cell,
        field,
        file,
        sheet,
        row,
        "无法解析为时间",
        dates::time_from_serial,
        dates::parse_time_text,
        Value::Time,
    )
}

/// 在表头中查找某统一字段的实际列号：候选同义词中恰好一个出现时返回该列号；
/// 均未出现时返回`None`；多个同义词同时出现视为结构异常。
pub(crate) fn resolve_synonym_column(
    header: &[String],
    synonyms: &[&str],
) -> Result<Option<u32>, String> {
    let mut found = synonyms.iter().filter_map(|&s| {
        header
            .iter()
            .position(|name| name == s)
            .map(|p| p as u32 + 1)
    });
    match (found.next(), found.next()) {
        (None, _) => Ok(None),
        (Some(col), None) => Ok(Some(col)),
        _ => Err(format!("字段候选名称 {synonyms:?} 在表头中出现多个匹配")),
    }
}

/// 由`Value::Date`与去“收款”前缀的单据号拼接匹配单据号；`doc_no_value`为空或`date`不是
/// `Value::Date`时留空（第9、10节共用）。
pub(crate) fn build_match_doc_no(date: &Value, doc_no_value: &str) -> String {
    if doc_no_value.is_empty() {
        return String::new();
    }
    match date {
        Value::Date(d) => doc_no::build_match_doc_no(*d, doc_no_value),
        _ => String::new(),
    }
}

/// `字段值 → 该值下全部候选（未去重）`，供`resolve`/`resolve_via`统一做歧义判定；
/// 用于按某个匹配键在另一数据源中查找唯一命中值的场景（如销售用券情况统计第10.10/
/// 10.12节、已上传数据与回款明细的匹配）。
pub(crate) type MultiValueIndex = HashMap<String, Vec<String>>;

pub(crate) enum PriorityOutcome {
    Unique(String),
    Ambiguous,
    NoHit,
}

/// 对一组候选值去重：恰好一个不同值→`Unique`；两个及以上→`Ambiguous`；没有候选→`NoHit`。
pub(crate) fn resolve(hits: Vec<String>) -> PriorityOutcome {
    let distinct: std::collections::HashSet<String> = hits.into_iter().collect();
    match distinct.len() {
        0 => PriorityOutcome::NoHit,
        1 => PriorityOutcome::Unique(distinct.into_iter().next().unwrap()),
        _ => PriorityOutcome::Ambiguous,
    }
}

/// 按`key`在索引中查找候选并去重；`key`为空或未在索引中出现时视为`NoHit`。
pub(crate) fn resolve_via(index: &MultiValueIndex, key: Option<&str>) -> PriorityOutcome {
    let Some(key) = key.filter(|k| !k.is_empty()) else {
        return PriorityOutcome::NoHit;
    };
    resolve(index.get(key).cloned().unwrap_or_default())
}

/// 按`key`查找唯一命中值；歧义或未命中均返回`None`，不得任选。
pub(crate) fn unique_hit(index: &MultiValueIndex, key: &str) -> Option<String> {
    match resolve_via(index, Some(key)) {
        PriorityOutcome::Unique(value) => Some(value),
        PriorityOutcome::Ambiguous | PriorityOutcome::NoHit => None,
    }
}

/// 记录并检查工作表内容指纹：同一指纹已属于另一个文件时判定为疑似重复导出并报错；
/// 否则记录该指纹归属的当前文件。
pub(crate) fn check_duplicate_fingerprint(
    fingerprints: &mut HashMap<String, String>,
    fingerprint: String,
    file_name: &str,
) -> Result<(), ProcessError> {
    if let Some(previous_file) = fingerprints.get(&fingerprint)
        && previous_file != file_name
    {
        return Err(ProcessError::Duplicate {
            detail: format!("{file_name} 与 {previous_file} 的工作表内容完全相同，疑似重复导出"),
        });
    }
    fingerprints
        .entry(fingerprint)
        .or_insert_with(|| file_name.to_string());
    Ok(())
}

/// 在`(路径, 排序键)`候选中选择键最大的唯一一项；键并列时判定为无法唯一确定并报错。
pub(crate) fn pick_unique_latest<K: Ord + Copy>(
    dated: Vec<(PathBuf, K)>,
    ambiguous_detail: &str,
) -> Result<PathBuf, ProcessError> {
    let max_key = dated.iter().map(|(_, key)| *key).max().unwrap();
    let mut latest: Vec<_> = dated
        .into_iter()
        .filter(|(_, key)| *key == max_key)
        .collect();
    if latest.len() > 1 {
        let names = latest
            .iter()
            .filter_map(|(path, _)| path.file_name())
            .map(|n| n.to_string_lossy().into_owned())
            .collect::<Vec<_>>()
            .join("、");
        return Err(ProcessError::Structure {
            file: names,
            sheet: String::new(),
            detail: ambiguous_detail.to_string(),
        });
    }
    Ok(latest.pop().unwrap().0)
}

pub mod coupons;
pub mod invoice;
pub mod receipts;
pub mod refund;
pub mod union_subsidy;
pub mod unionpay;
pub mod uploaded;

static REGISTRY: [&(dyn Job + Sync); 9] = [
    &invoice::InvoiceJob,
    &union_subsidy::UnionSubsidyJob,
    &uploaded::UPLOADED_DIGITAL,
    &uploaded::UPLOADED_APPLIANCE,
    &unionpay::UnionPayJob,
    &refund::REFUND_APPLIANCE,
    &refund::REFUND_DIGITAL,
    &receipts::ReceiptsJob,
    &coupons::CouponsJob,
];

/// 按菜单编号 1–9 排列的全部任务；`app::runner`据此驱动单类或批量执行。
pub fn registry() -> &'static [&'static (dyn Job + Sync)] {
    &REGISTRY
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::str::FromStr;

    #[test]
    fn registry_is_ordered_by_menu_number_one_through_nine() {
        let numbers: Vec<u32> = registry().iter().map(|job| job.category() as u32).collect();
        assert_eq!(numbers, (1..=9).collect::<Vec<_>>());
    }

    #[test]
    fn cell_amount_parses_numeric_text() {
        assert_eq!(
            cell_amount(&RawCell::Text("1000.00".to_string())),
            Ok(Some(Decimal::from_str("1000.00").unwrap()))
        );
    }

    #[test]
    fn cell_amount_blank_text_is_none() {
        assert_eq!(cell_amount(&RawCell::Text("   ".to_string())), Ok(None));
        assert_eq!(cell_amount(&RawCell::Empty), Ok(None));
    }

    #[test]
    fn cell_amount_rejects_non_numeric_text() {
        assert!(cell_amount(&RawCell::Text("不是数字".to_string())).is_err());
    }

    #[test]
    fn cell_text_preserves_whitespace_only_text() {
        assert_eq!(
            cell_text(&RawCell::Text("  ".to_string())),
            Ok("  ".to_string())
        );
    }

    #[test]
    fn text_value_and_amount_value_map_empty_to_value_empty() {
        assert_eq!(text_value(String::new()), Value::Empty);
        assert_eq!(text_value("x".to_string()), Value::Text("x".to_string()));
        assert_eq!(amount_value(None), Value::Empty);
    }
}
