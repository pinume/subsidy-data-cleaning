use std::collections::{HashMap, HashSet};
use std::path::Path;
use std::sync::LazyLock;

use regex::Regex;
use rust_decimal::Decimal;
use rust_decimal::prelude::FromPrimitive;

use crate::io::xlsx_reader::{RawCell, SheetGrid, open_sheets};
use crate::model::{Column, ColumnType, Fill, ProcessError, Row, Table, Value};
use crate::utils::{numbers, text};

use super::{
    Category, Job, MultiValueIndex, PriorityOutcome, build_match_doc_no, cell_text, data_error,
    parse_date_field, resolve, resolve_via, text_value, unique_hit,
};
use super::{invoice, receipts, unionpay, uploaded};

const FILE_NAME: &str = "销售用券情况统计.xlsx";
const TITLE: &str = "销售用券情况统计";

const SOURCE_HEADERS: [&str; 28] = [
    "供应商",
    "原始供应商",
    "单据号",
    "单据日期",
    "收款员",
    "商品名称",
    "商品简码",
    "品牌",
    "销售部门",
    "销售员",
    "业绩成本",
    "业绩利润",
    "价格类型名称",
    "库存类型",
    "财务大类",
    "顾客姓名",
    "备注",
    "明细摘要",
    "收款日期",
    "销售成本",
    "不含券收入",
    "其它",
    "本期尾款",
    "销售单价",
    "含券收入",
    "2026家电国补（计入收入）",
    "2026数码国补（计入收入）",
    "合计",
];

// 源表列位置（1 基）。
const COL_DOC_NO: u32 = 3;
const COL_DOC_DATE: u32 = 4;
const COL_PRODUCT_NAME: u32 = 6;
const COL_BRAND: u32 = 8;
const COL_FINANCE_CATEGORY: u32 = 15;
const COL_SUMMARY: u32 = 18;
const COL_TOTAL: u32 = 28;

const TRIGGER_LETTERS: [char; 10] = ['N', 'n', 'W', 'w', 'M', 'm', 'H', 'h', 'B', 'b'];

struct CouponRecord {
    doc_no: String,
    doc_date: Value,
    product_name: String,
    brand: String,
    finance_category: String,
    subsidy: Decimal,
    summary: String,
}

pub struct CouponsJob;

impl Job for CouponsJob {
    fn category(&self) -> Category {
        Category::Coupons
    }

    fn title(&self) -> &'static str {
        "销售用券情况统计"
    }

    fn output_stem(&self) -> &'static str {
        "销售用券情况统计"
    }

    fn run(&self, input_dir: &Path) -> Result<Table, ProcessError> {
        // 10.1 节：先校验本任务自身输入文件的结构，再进入依赖外部数据源的字段生成步骤，
        // 确保输入文件本身的问题不会被外部依赖缺失的错误提示掩盖。
        let path = input_dir.join(FILE_NAME);
        if !path.is_file() {
            return Err(ProcessError::NoInput {
                pattern: FILE_NAME.to_string(),
            });
        }

        let sheets = open_sheets(&path)?;
        if sheets.len() != 1 {
            return Err(ProcessError::Structure {
                file: FILE_NAME.to_string(),
                sheet: String::new(),
                detail: format!("工作表数量异常：应为 1 个，实际为 {} 个", sheets.len()),
            });
        }
        let sheet = &sheets[0];
        let sheet_name = sheet.name().to_string();

        let title = sheet.cell(1, 1).to_string();
        if title != TITLE {
            return Err(ProcessError::Structure {
                file: FILE_NAME.to_string(),
                sheet: sheet_name,
                detail: format!("第1行第1列应为“{TITLE}”，实际为“{title}”"),
            });
        }

        let header = sheet.row_texts(2);
        if header.iter().map(String::as_str).collect::<Vec<_>>() != SOURCE_HEADERS {
            return Err(ProcessError::Structure {
                file: FILE_NAME.to_string(),
                sheet: sheet_name,
                detail: "第2行表头与规定的28个字段不一致".to_string(),
            });
        }

        let last_row = sheet.last_value_row().unwrap_or(2);
        let total_marker = sheet.cell(last_row, 1).to_string();
        if total_marker != "合计" {
            return Err(ProcessError::Structure {
                file: FILE_NAME.to_string(),
                sheet: sheet_name,
                detail: format!("最后一个实际有值行第1列应为“合计”，实际为“{total_marker}”"),
            });
        }

        // 10.6.1：权威校验集缺失、结构异常或无法完整读取时必须停止，不得绕过校验；
        // 直接复用 unionpay::load_records 的全部校验（文件发现、表头、首尾结构、重复导出）。
        let authority = build_authority(input_dir)?;
        // 数电发票号码：缺失或结构异常时同样必须停止，直接复用 invoice::load_records
        // 的全部校验（最新文件选择、表头、逐行解析）。
        let invoice_index = build_invoice_index(input_dir)?;
        // 备注：缺失或结构异常时同样必须停止，直接复用 receipts::load_records 的全部
        // 校验（表头、末行合计结构）及其已算好的第 9.5 节三阶段备注。
        let receipts_index = build_receipts_index(input_dir)?;
        // 备注兜底：缺失或结构异常时同样必须停止，直接复用已上传家电电脑/已上传数码
        // 两个 Job 各自的全部校验（文件发现、表头、第26列起字段解析、重复UUID）。
        let (uploaded_by_reference, uploaded_by_invoice_no) = build_uploaded_index(input_dir)?;

        let mut rows = Vec::new();
        for row in 3..last_row {
            let record = read_row(sheet, row, FILE_NAME, &sheet_name)?;
            rows.push(to_row(
                record,
                &authority,
                &invoice_index,
                &receipts_index,
                &uploaded_by_reference,
                &uploaded_by_invoice_no,
            ));
        }

        // 命中收款单统计备注的记录沉底并保持粉色；其余保持原相对顺序（各自稳定分区）。
        let (mut top, bottom): (Vec<_>, Vec<_>) =
            rows.into_iter().partition(|r| r.fill != Some(Fill::Pink));
        top.extend(bottom);

        Ok(Table {
            columns: output_columns(),
            rows: top,
        })
    }
}

fn build_authority(input_dir: &Path) -> Result<HashSet<String>, ProcessError> {
    let records = unionpay::load_records(input_dir)?;
    Ok(records
        .iter()
        .map(|r| r.retrieval_no.as_str())
        .filter(|v| is_valid_ref_no(v))
        .map(str::to_owned)
        .collect())
}

fn is_valid_ref_no(value: &str) -> bool {
    value.len() == 12
        && value.as_bytes()[11] == b'N'
        && value.as_bytes()[..11].iter().all(u8::is_ascii_digit)
}

/// 按`匹配单据号`汇总发票明细全部非空`数电发票号码`（未去重）；歧义判定见`to_row`
/// 中复用的`resolve`（与 10.6.2 节"命中权威值需唯一"同一原则，不得任选）。
fn build_invoice_index(input_dir: &Path) -> Result<MultiValueIndex, ProcessError> {
    let records = invoice::load_records(input_dir)?;
    let mut grouped: MultiValueIndex = HashMap::new();
    for record in &records {
        let Value::Text(match_doc_no) = &record.match_doc_no else {
            continue;
        };
        if record.invoice_no.is_empty() {
            continue;
        }
        grouped
            .entry(match_doc_no.clone())
            .or_default()
            .push(record.invoice_no.clone());
    }
    Ok(grouped)
}

/// 按`匹配单据号`汇总收款单统计全部非空`备注`（未去重）；歧义判定同样复用`resolve`。
fn build_receipts_index(input_dir: &Path) -> Result<MultiValueIndex, ProcessError> {
    let records = receipts::load_records(input_dir)?;
    let mut grouped: MultiValueIndex = HashMap::new();
    for record in &records {
        if record.match_doc_no.is_empty() || record.remark.is_empty() {
            continue;
        }
        grouped
            .entry(record.match_doc_no.clone())
            .or_default()
            .push(record.remark.clone());
    }
    Ok(grouped)
}

// 已上传家电电脑/已上传数码前25列固定结构中的字段位置（0 基，对齐 Table.rows[].values）。
const UPLOADED_COL_REFERENCE: usize = 6; // 检索参考号
const UPLOADED_COL_STATUS: usize = 8; // 状态
const UPLOADED_COL_INVOICE_NO: usize = 19; // 发票号码

/// 按`检索参考号`和`发票号码`分别汇总已上传家电电脑、已上传数码合并后的全部非空`状态`
/// （未去重）；两个数据组共用同一对索引，不按财务大类等字段区分数据组。
fn build_uploaded_index(
    input_dir: &Path,
) -> Result<(MultiValueIndex, MultiValueIndex), ProcessError> {
    let mut by_reference: MultiValueIndex = HashMap::new();
    let mut by_invoice_no: MultiValueIndex = HashMap::new();

    for job in [&uploaded::UPLOADED_APPLIANCE, &uploaded::UPLOADED_DIGITAL] {
        let table = job.run(input_dir)?;
        for row in &table.rows {
            let Value::Text(status) = &row.values[UPLOADED_COL_STATUS] else {
                continue;
            };
            if let Value::Text(reference) = &row.values[UPLOADED_COL_REFERENCE] {
                by_reference
                    .entry(reference.clone())
                    .or_default()
                    .push(status.clone());
            }
            if let Value::Text(invoice_no) = &row.values[UPLOADED_COL_INVOICE_NO] {
                by_invoice_no
                    .entry(invoice_no.clone())
                    .or_default()
                    .push(status.clone());
            }
        }
    }
    Ok((by_reference, by_invoice_no))
}

/// 按`参考号`（主键）优先、`数电发票号码`（次键）兜底，在已上传数据索引中查找唯一命中的
/// `状态`；仅当主键未命中（键为空或索引查无）时才尝试次键，歧义则立即留空、不再降级。
fn uploaded_status(
    by_reference: &MultiValueIndex,
    by_invoice_no: &MultiValueIndex,
    reference: Option<&str>,
    invoice_no: Option<&str>,
) -> Option<String> {
    match resolve_via(by_reference, reference) {
        PriorityOutcome::Unique(value) => return Some(value),
        PriorityOutcome::Ambiguous => return None,
        PriorityOutcome::NoHit => {}
    }
    match resolve_via(by_invoice_no, invoice_no) {
        PriorityOutcome::Unique(value) => Some(value),
        PriorityOutcome::Ambiguous | PriorityOutcome::NoHit => None,
    }
}

/// 补贴额：取源字段`合计`，按 10.7 节尾差规则统一为两位小数；为空或超出容差时终止。
fn parse_subsidy(
    cell: &RawCell,
    file: &str,
    sheet: &str,
    row: u32,
) -> Result<Decimal, ProcessError> {
    let raw: Option<Decimal> = match cell {
        RawCell::Text(t) if t.trim().is_empty() => None,
        RawCell::Int(n) => Some(Decimal::from(*n)),
        RawCell::Float(f) => Decimal::from_f64(*f),
        RawCell::Text(t) => t.trim().parse().ok(),
        _ => None,
    };
    let Some(raw) = raw else {
        return Err(data_error(
            file,
            sheet,
            row,
            "合计",
            cell.to_string(),
            "补贴额为空或无法解析为数值".to_string(),
        ));
    };
    numbers::to_cents(raw).ok_or_else(|| {
        data_error(
            file,
            sheet,
            row,
            "合计",
            cell.to_string(),
            "金额与最接近的两位小数之差超出0.000001元容差".to_string(),
        )
    })
}

fn read_row(
    sheet: &SheetGrid,
    row: u32,
    file: &str,
    sheet_name: &str,
) -> Result<CouponRecord, ProcessError> {
    let text_at = |col: u32, field: &'static str| -> Result<String, ProcessError> {
        let cell = sheet.cell(row, col);
        cell_text(&cell)
            .map_err(|detail| data_error(file, sheet_name, row, field, cell.to_string(), detail))
    };

    Ok(CouponRecord {
        doc_no: text_at(COL_DOC_NO, "单据号")?,
        doc_date: parse_date_field(
            &sheet.cell(row, COL_DOC_DATE),
            "单据日期",
            file,
            sheet_name,
            row,
        )?,
        product_name: text_at(COL_PRODUCT_NAME, "商品名称")?,
        brand: text_at(COL_BRAND, "品牌")?,
        finance_category: text_at(COL_FINANCE_CATEGORY, "财务大类")?,
        subsidy: parse_subsidy(&sheet.cell(row, COL_TOTAL), file, sheet_name, row)?,
        summary: text_at(COL_SUMMARY, "明细摘要")?,
    })
}

fn quantity_of(subsidy: Decimal) -> i64 {
    use std::cmp::Ordering;
    match subsidy.cmp(&Decimal::ZERO) {
        Ordering::Greater => 1,
        Ordering::Equal => 0,
        Ordering::Less => -1,
    }
}

fn to_row(
    record: CouponRecord,
    authority: &HashSet<String>,
    invoice_index: &MultiValueIndex,
    receipts_index: &MultiValueIndex,
    uploaded_by_reference: &MultiValueIndex,
    uploaded_by_invoice_no: &MultiValueIndex,
) -> Row {
    let quantity = quantity_of(record.subsidy);
    let match_doc_no = build_match_doc_no(&record.doc_date, &record.doc_no);
    let reference = extract_reference(&record.summary, authority);
    let invoice_no = unique_hit(invoice_index, &match_doc_no);
    let receipts_hit = unique_hit(receipts_index, &match_doc_no);
    let fill = receipts_hit.is_some().then_some(Fill::Pink);
    let remark = receipts_hit.or_else(|| {
        uploaded_status(
            uploaded_by_reference,
            uploaded_by_invoice_no,
            reference.as_deref(),
            invoice_no.as_deref(),
        )
    });
    // 10.12.3 节：两阶段均未命中（含歧义、空键）时，备注固定填“未上传”，不沉底不填色；
    // 空值从不参与前两阶段的匹配，这里只是给最终仍为空的结果一个统一的文本标记。
    let remark_text = remark.unwrap_or_else(|| "未上传".to_string());

    let values = vec![
        text_value(record.doc_no),
        record.doc_date,
        text_value(record.product_name),
        text_value(normalize_brand(record.brand)),
        text_value(normalize_finance_category(record.finance_category)),
        Value::Decimal(record.subsidy),
        Value::Integer(quantity),
        reference.map_or(Value::Empty, Value::Text),
        invoice_no.map_or(Value::Empty, Value::Text),
        text_value(match_doc_no),
        Value::Text(remark_text),
    ];
    Row { values, fill }
}

/// 10.11 节品牌归并：仅替换列出的原值，其余原样保留。
fn normalize_brand(value: String) -> String {
    match value.as_str() {
        "COLMO厨热JX" | "美的厨热JX" | "东芝JX" | "华凌" | "小天鹅" | "COLMO" => {
            "美的".to_string()
        }
        "卡萨帝" | "统帅" => "海尔".to_string(),
        "晶弘" => "格力".to_string(),
        _ => value,
    }
}

/// 10.11 节财务大类归并：仅替换列出的原值，其余原样保留。
fn normalize_finance_category(value: String) -> String {
    match value.as_str() {
        "国产彩电" | "进口彩电" => "彩电".to_string(),
        "新业务类" => "数码".to_string(),
        _ => value,
    }
}

fn output_columns() -> Vec<Column> {
    vec![
        Column {
            name: "单据号",
            ty: ColumnType::Text,
        },
        Column {
            name: "单据日期",
            ty: ColumnType::Date,
        },
        Column {
            name: "商品名称",
            ty: ColumnType::Text,
        },
        Column {
            name: "品牌",
            ty: ColumnType::Text,
        },
        Column {
            name: "财务大类",
            ty: ColumnType::Text,
        },
        Column {
            name: "补贴额",
            ty: ColumnType::Decimal(crate::model::DecimalScale::Two),
        },
        Column {
            name: "数量",
            ty: ColumnType::Integer,
        },
        Column {
            name: "参考号",
            ty: ColumnType::Text,
        },
        Column {
            name: "数电发票号码",
            ty: ColumnType::Text,
        },
        Column {
            name: "匹配单据号",
            ty: ColumnType::Text,
        },
        Column {
            name: "备注",
            ty: ColumnType::Text,
        },
    ]
}

// ---------------------------------------------------------------------------
// 10.6 节：参考号提取（四级优先级，命中即停，同级内先去重再判定）。
// ---------------------------------------------------------------------------

fn extract_reference(summary: &str, authority: &HashSet<String>) -> Option<String> {
    if summary.trim().is_empty() {
        return None;
    }

    match resolve(as_is_hits(summary, authority)) {
        PriorityOutcome::Unique(value) => return Some(value),
        PriorityOutcome::Ambiguous => return None,
        PriorityOutcome::NoHit => {}
    }

    let digit_candidates = digit_candidates(summary);

    match resolve(zero_edit_hits(&digit_candidates, authority)) {
        PriorityOutcome::Unique(value) => return Some(value),
        PriorityOutcome::Ambiguous => return None,
        PriorityOutcome::NoHit => {}
    }

    match resolve(single_edit_hits(&digit_candidates, authority)) {
        PriorityOutcome::Unique(value) => Some(value),
        PriorityOutcome::Ambiguous | PriorityOutcome::NoHit => None,
    }
}

fn reference_pattern() -> &'static Regex {
    static RE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"[0-9]{11}N").unwrap());
    &RE
}

/// 第 1 级：原样提取完整`[0-9]{11}N`片段，前后不得紧邻数字或字母。
fn as_is_hits(summary: &str, authority: &HashSet<String>) -> Vec<String> {
    reference_pattern()
        .find_iter(summary)
        .filter(|m| text::has_isolated_boundaries(summary, m.start(), m.end()))
        .map(|m| m.as_str().to_string())
        .filter(|candidate| authority.contains(candidate))
        .collect()
}

/// 字节级扫描，返回所有连续 ASCII 数字片段的字节范围（Chinese 字符不会被误判）。
fn maximal_digit_runs(text: &str) -> Vec<(usize, usize)> {
    let bytes = text.as_bytes();
    let mut runs = Vec::new();
    let mut start = None;
    for (i, &b) in bytes.iter().enumerate() {
        if b.is_ascii_digit() {
            start.get_or_insert(i);
        } else if let Some(s) = start.take() {
            runs.push((s, i));
        }
    }
    if let Some(s) = start {
        runs.push((s, bytes.len()));
    }
    runs
}

/// 第 2 级候选：长度 10-12 的完整连续数字串；若摘要含触发字母，额外把全部数字顺序拼接为一个候选。
fn digit_candidates(summary: &str) -> Vec<String> {
    let mut candidates = Vec::new();

    for (start, end) in maximal_digit_runs(summary) {
        if (10..=12).contains(&(end - start)) {
            candidates.push(summary[start..end].to_string());
        }
    }

    if summary.chars().any(|c| TRIGGER_LETTERS.contains(&c)) {
        let all_digits: String = summary.chars().filter(char::is_ascii_digit).collect();
        if (10..=12).contains(&all_digits.len()) {
            candidates.push(all_digits);
        }
    }

    candidates
}

/// 第 3 级：仅对 11 位候选补上大写`N`后与权威校验集比对，不做任何数字改动。
fn zero_edit_hits(candidates: &[String], authority: &HashSet<String>) -> Vec<String> {
    candidates
        .iter()
        .filter(|c| c.len() == 11)
        .map(|c| format!("{c}N"))
        .filter(|candidate| authority.contains(candidate))
        .collect()
}

/// 第 4 级：对 10-12 位候选执行恰好一次插入/替换/删除，纠正为 11 位后补`N`比对。
fn single_edit_hits(candidates: &[String], authority: &HashSet<String>) -> Vec<String> {
    let mut hits = Vec::new();
    for candidate in candidates {
        for variant in single_edit_variants(candidate) {
            let with_suffix = format!("{variant}N");
            if authority.contains(&with_suffix) {
                hits.push(with_suffix);
            }
        }
    }
    hits
}

fn single_edit_variants(candidate: &str) -> Vec<String> {
    let digits = candidate.as_bytes();
    let mut variants = Vec::new();
    match digits.len() {
        10 => {
            for pos in 0..=digits.len() {
                for d in b'0'..=b'9' {
                    let mut v = digits.to_vec();
                    v.insert(pos, d);
                    variants.push(String::from_utf8(v).unwrap());
                }
            }
        }
        11 => {
            for pos in 0..digits.len() {
                for d in b'0'..=b'9' {
                    if d != digits[pos] {
                        let mut v = digits.to_vec();
                        v[pos] = d;
                        variants.push(String::from_utf8(v).unwrap());
                    }
                }
            }
        }
        12 => {
            for pos in 0..digits.len() {
                let mut v = digits.to_vec();
                v.remove(pos);
                variants.push(String::from_utf8(v).unwrap());
            }
        }
        _ => {}
    }
    variants
}

#[cfg(test)]
mod tests {
    use rust_xlsxwriter::Workbook;

    use super::*;
    use crate::jobs::{invoice, receipts, refund, unionpay, uploaded};

    const MERCHANT_DIGITAL: &str = "89813014812B06R";
    const MERCHANT_APPLIANCE: &str = "89813015722APT1";
    use crate::test_support::unique_temp_path;

    fn authority_of(values: &[&str]) -> HashSet<String> {
        values.iter().map(ToString::to_string).collect()
    }

    #[test]
    fn validates_reference_number_shape() {
        assert!(is_valid_ref_no("16867252734N"));
        assert!(!is_valid_ref_no("16867252734W")); // 结尾不是大写 N
        assert!(!is_valid_ref_no("1686725273N")); // 只有 10 位数字
        assert!(!is_valid_ref_no("16867252734n")); // 小写 n 不算
    }

    #[test]
    fn normalizes_listed_brands_and_keeps_others() {
        for brand in [
            "COLMO厨热JX",
            "美的厨热JX",
            "东芝JX",
            "华凌",
            "小天鹅",
            "COLMO",
        ] {
            assert_eq!(normalize_brand(brand.to_string()), "美的");
        }
        for brand in ["卡萨帝", "统帅"] {
            assert_eq!(normalize_brand(brand.to_string()), "海尔");
        }
        assert_eq!(normalize_brand("晶弘".to_string()), "格力");
        assert_eq!(normalize_brand("海信".to_string()), "海信"); // 未列出的品牌原样保留
    }

    #[test]
    fn normalizes_listed_finance_categories_and_keeps_others() {
        assert_eq!(normalize_finance_category("国产彩电".to_string()), "彩电");
        assert_eq!(normalize_finance_category("进口彩电".to_string()), "彩电");
        assert_eq!(normalize_finance_category("新业务类".to_string()), "数码");
        assert_eq!(normalize_finance_category("空调".to_string()), "空调"); // 未列出的类别原样保留
    }

    #[test]
    fn quantity_sign_matches_subsidy() {
        assert_eq!(quantity_of("10.00".parse().unwrap()), 1);
        assert_eq!(quantity_of("0.00".parse().unwrap()), 0);
        assert_eq!(quantity_of("-10.00".parse().unwrap()), -1);
    }

    #[test]
    fn as_is_hits_requires_isolated_match_in_authority() {
        let authority = authority_of(&["16867252734N"]);
        assert_eq!(
            as_is_hits("摘要16867252734N末尾", &authority),
            vec!["16867252734N".to_string()]
        );
        // 紧邻的后续数字使其成为更长数字串的一部分，不得原样提取。
        assert!(as_is_hits("摘要16867252734N9末尾", &authority).is_empty());
        // 格式正确但不在权威校验集中的值不得输出。
        assert!(as_is_hits("摘要99999999999N末尾", &authority).is_empty());
    }

    #[test]
    fn extracts_reference_via_as_is_match() {
        let authority = authority_of(&["16867252734N"]);
        assert_eq!(
            extract_reference("参考号：16867252734N。", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn falls_back_to_zero_edit_when_n_suffix_missing() {
        let authority = authority_of(&["16867252734N"]);
        // 只有 11 位数字，缺少末尾大写 N；原样提取无结果，第三级补上 N 后命中。
        assert_eq!(
            extract_reference("单号16867252734完成", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn trigger_letter_allows_digit_concatenation_across_gaps() {
        let authority = authority_of(&["16867252734N"]);
        // 数字被空格分割为两段（6 位+5 位，均不在 10-12 位范围内），
        // 但摘要含触发字母 N，可将全部数字按原顺序拼接为一个候选。
        assert_eq!(
            extract_reference("订单168672 52734N附言", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn single_digit_insertion_recovers_reference() {
        let authority = authority_of(&["16867252734N"]);
        // 缺少末位数字 4（10 位），第四级允许一次插入。
        assert_eq!(
            extract_reference("单号1686725273结清", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn single_digit_substitution_recovers_reference() {
        let authority = authority_of(&["16867252734N"]);
        // 末位数字错为 9（应为 4），第四级允许一次替换。
        assert_eq!(
            extract_reference("单号16867252739结清", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn single_digit_deletion_recovers_reference() {
        let authority = authority_of(&["16867252734N"]);
        // 多出一位数字（12 位），第四级允许一次删除。
        assert_eq!(
            extract_reference("单号168672527340结清", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn two_edits_are_not_attempted() {
        let authority = authority_of(&["16867252734N"]);
        // 两位数字都错误，超出“至多一次改动”的范围，不得纠正。
        assert_eq!(extract_reference("单号16867252799结清", &authority), None);
    }

    #[test]
    fn ambiguous_hits_at_same_priority_leave_reference_blank() {
        let authority = authority_of(&["11111111111N", "22222222222N"]);
        assert_eq!(
            extract_reference("含11111111111N及22222222222N两个编号", &authority),
            None
        );
    }

    #[test]
    fn blank_or_unrelated_summary_yields_none() {
        let authority = authority_of(&["16867252734N"]);
        assert_eq!(extract_reference("", &authority), None);
        assert_eq!(extract_reference("预售", &authority), None);
    }

    // --- 集成测试：构造一份门店银联样本建立权威校验集，再跑完整 CouponsJob。 ---

    fn write_unionpay_fixture(dir: &Path, retrieval_no: &str) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("对账数据").unwrap();
        sheet.write_string(0, 0, "汇总").unwrap();
        for (col, header) in unionpay::HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        let row: [&str; 26] = [
            "20260914",
            "2026-09-14 10:18:09",
            "T001",
            "消费",
            "622***1234",
            "100.00",
            "100.00",
            "1.00",
            "0.50",
            "0.50",
            "SN0001",
            retrieval_no,
            "借记卡",
            "工商银行",
            "89813014812B06R",
            "某门店",
            "门店简",
            "ORDER1",
            "AC0001",
            "云闪付",
            "分店A",
            "0.00",
            "0.00",
            "备注文字",
            "",
            "buyer001",
        ];
        for (col, value) in row.iter().enumerate() {
            sheet.write_string(2, col as u16, *value).unwrap();
        }
        sheet.write_string(3, 0, unionpay::D1_NOTICE).unwrap();
        workbook
            .save(dir.join("89813014812B06R_MX_20260914101809_1.xlsx"))
            .unwrap();
    }

    /// 构造一份最小合法的发票明细样本（单工作表、30 字段表头），供数电发票号码匹配测试；
    /// `rows`为`(备注信息, 数电发票号码)`对，其余字段使用固定合法占位值。
    fn write_invoice_fixture(dir: &Path, rows: &[(&str, &str)]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("发票_20260914").unwrap();
        for (col, header) in invoice::SOURCE_HEADERS.iter().enumerate() {
            sheet.write_string(5, col as u16, *header).unwrap();
        }
        for (index, (remark, invoice_no)) in rows.iter().enumerate() {
            let row = (6 + index) as u32;
            let values: [&str; 30] = [
                "ORDER001",
                "2026-09-14 09:00:00",
                "2026-09-14 10:00:00",
                "蓝票",
                "数电发票",
                "正常发票",
                "",
                "",
                invoice_no,
                "张三",
                "",
                "",
                "",
                "",
                "",
                "商品甲",
                "100.00",
                "13%",
                "88.50",
                "11.50",
                remark,
                "总店",
                "自动开票",
                "开票员甲",
                "收款人甲",
                "复核人甲",
                "",
                "开票完成",
                "操作人甲",
                "已打印",
            ];
            for (col, value) in values.iter().enumerate() {
                sheet.write_string(row, col as u16, *value).unwrap();
            }
        }
        workbook.save(dir.join("发票_20260914.xlsx")).unwrap();
    }

    /// 构造一份最小合法的收款单统计样本（单工作表、60 字段表头、标题+表头+合计结构），
    /// 供备注匹配测试；`rows`为`(日期, 单据号, 销售类别, 原票号)`四元组，其余字段留空。
    /// `销售类别="退货"`可直接得到非空初始备注`退货-退单`，无需再构造跨阶段关联。
    fn write_receipts_fixture(dir: &Path, rows: &[(&str, &str, &str, &str)]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, "标题").unwrap();
        for (col, header) in receipts::SOURCE_HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        for (index, (date, doc_no, sale_category, original_ticket_no)) in rows.iter().enumerate() {
            let row = (2 + index) as u32;
            sheet.write_string(row, 0, *date).unwrap(); // 日期
            sheet.write_string(row, 2, *doc_no).unwrap(); // 单据号
            sheet.write_string(row, 8, *sale_category).unwrap(); // 销售类别
            sheet.write_string(row, 40, *original_ticket_no).unwrap(); // 原票号
        }
        let total_row = (2 + rows.len()) as u32;
        sheet.write_string(total_row, 0, "合计").unwrap();
        workbook.save(dir.join("收款单统计.xlsx")).unwrap();
    }

    /// 构造一份最小合法的已上传数据样本（单工作表、58 字段表头：前25列固定 + 对应
    /// 数据组的尾部字段名），供备注兜底匹配测试；`rows`为
    /// `(实时清分UUID, 检索参考号, 状态, 发票号码)`四元组，其余字段留空。
    fn write_uploaded_fixture(
        dir: &Path,
        merchant_no: &str,
        tail: &[uploaded::TailField],
        rows: &[(&str, &str, &str, &str)],
    ) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, "标题").unwrap();
        for (col, header) in uploaded::FRONT_HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        for (col, field) in tail.iter().enumerate() {
            sheet
                .write_string(
                    1,
                    (uploaded::FRONT_HEADERS.len() + col) as u16,
                    field.synonyms[0],
                )
                .unwrap();
        }
        for (index, (uuid, reference, status, invoice_no)) in rows.iter().enumerate() {
            let row = (2 + index) as u32;
            sheet.write_string(row, 0, *uuid).unwrap(); // 实时清分UUID
            sheet.write_string(row, 1, merchant_no).unwrap(); // 商户号（须与文件名一致）
            sheet.write_string(row, 6, *reference).unwrap(); // 检索参考号
            sheet.write_string(row, 8, *status).unwrap(); // 状态
            sheet.write_string(row, 19, *invoice_no).unwrap(); // 发票号码
        }
        workbook
            .save(dir.join(format!("MER_{merchant_no}_20260914101809_yjhx.xlsx")))
            .unwrap();
    }

    /// 写入一份没有明细的最小合法回款明细样本（家电电脑或数码），仅用于满足已上传数据
    /// 内部匹配补贴金额时的前置依赖存在；24列表头取回款明细的统一确认名称，对两个数据组
    /// 的必需字段判定均有效。
    fn write_empty_refund_fixture(dir: &Path, filename: &str) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        for (col, header) in refund::OUTPUT_FIELDS.iter().enumerate() {
            sheet.write_string(0, col as u16, *header).unwrap();
        }
        workbook.save(dir.join(filename)).unwrap();
    }

    /// 两个已上传数据组各写入一份没有明细的最小合法样本，并同时写入两份空回款明细样本
    /// （已上传数据内部会按订单号/检索参考号/发票号码匹配回款明细的补贴金额），仅用于
    /// 满足前置依赖存在。
    fn write_empty_uploaded_fixtures(dir: &Path) {
        write_empty_refund_fixture(dir, "2026年以旧换新补贴明细.xlsx");
        write_empty_refund_fixture(dir, "2026年数码补贴明细.xlsx");
        write_uploaded_fixture(dir, MERCHANT_APPLIANCE, &uploaded::APPLIANCE_TAIL, &[]);
        write_uploaded_fixture(dir, MERCHANT_DIGITAL, &uploaded::DIGITAL_TAIL, &[]);
    }

    fn write_coupons_workbook(dir: &Path, rows: &[(&str, &str, &str, &str, &str, &str, &str)]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, TITLE).unwrap();
        for (col, header) in SOURCE_HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        for (index, (doc_no, doc_date, product_name, brand, finance_category, summary, total)) in
            rows.iter().enumerate()
        {
            let row = (2 + index) as u32;
            sheet
                .write_string(row, (COL_DOC_NO - 1) as u16, *doc_no)
                .unwrap();
            sheet
                .write_string(row, (COL_DOC_DATE - 1) as u16, *doc_date)
                .unwrap();
            sheet
                .write_string(row, (COL_PRODUCT_NAME - 1) as u16, *product_name)
                .unwrap();
            sheet
                .write_string(row, (COL_BRAND - 1) as u16, *brand)
                .unwrap();
            sheet
                .write_string(row, (COL_FINANCE_CATEGORY - 1) as u16, *finance_category)
                .unwrap();
            sheet
                .write_string(row, (COL_SUMMARY - 1) as u16, *summary)
                .unwrap();
            sheet
                .write_number(row, (COL_TOTAL - 1) as u16, total.parse::<f64>().unwrap())
                .unwrap();
        }
        sheet
            .write_string((2 + rows.len()) as u32, 0, "合计")
            .unwrap();
        workbook.save(dir.join(FILE_NAME)).unwrap();
    }

    #[test]
    fn end_to_end_extracts_reference_and_generates_match_doc_no() {
        let dir = unique_temp_path("coupons-happy-path");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        // 第一条发票明细的匹配单据号与销售明细第一行一致，可用于命中数电发票号码；
        // 第二行销售明细的匹配单据号在发票明细中没有对应记录，应留空。
        write_invoice_fixture(
            &dir,
            &[(
                "销售日期:2026-08-29 单据号:收款ZFFX000003",
                "24312000000000000009",
            )],
        );
        // 收款单统计中没有任何匹配单据号能命中这两行，备注列应全部留空、不触发沉底。
        write_receipts_fixture(&dir, &[]);
        write_empty_uploaded_fixtures(&dir);
        write_coupons_workbook(
            &dir,
            &[
                (
                    "收款ZFFX000003",
                    "2026-08-29",
                    "商品甲",
                    "品牌甲",
                    "家电",
                    "参考号：16867252734N",
                    "299.85",
                ),
                (
                    "ZFFX000004",
                    "2026-08-30",
                    "商品乙",
                    "品牌乙",
                    "数码",
                    "无编号信息",
                    "-50.00",
                ),
            ],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(table.columns.len(), 11);
        assert_eq!(table.rows.len(), 2);

        // 第一行：参考号原样命中，补贴额为正，数量为 1，匹配单据号保留“收款”前缀已剥离后拼接，
        // 并按匹配单据号在发票明细中唯一命中数电发票号码。
        assert_eq!(
            table.rows[0].values[0],
            Value::Text("收款ZFFX000003".to_string())
        );
        assert_eq!(
            table.rows[0].values[5],
            Value::Decimal("299.85".parse().unwrap())
        );
        assert_eq!(table.rows[0].values[6], Value::Integer(1));
        assert_eq!(
            table.rows[0].values[7],
            Value::Text("16867252734N".to_string())
        );
        assert_eq!(
            table.rows[0].values[8],
            Value::Text("24312000000000000009".to_string())
        );
        assert_eq!(
            table.rows[0].values[9],
            Value::Text("260829ZFFX000003".to_string())
        );
        // 收款单统计及已上传数据（本用例为空样本）均无命中，备注按10.12.3节固定为“未上传”。
        assert_eq!(table.rows[0].values[10], Value::Text("未上传".to_string()));
        assert_eq!(table.rows[0].fill, None); // 未命中备注，不沉底不填色

        // 第二行：摘要中无编号，参考号留空；补贴额为负，数量为 -1；
        // 匹配单据号在发票明细中无对应记录，数电发票号码留空。
        assert_eq!(table.rows[1].values[6], Value::Integer(-1));
        assert_eq!(table.rows[1].values[7], Value::Empty);
        assert_eq!(table.rows[1].values[8], Value::Empty);
        assert_eq!(
            table.rows[1].values[9],
            Value::Text("260830ZFFX000004".to_string())
        );
        assert_eq!(table.rows[1].values[10], Value::Text("未上传".to_string()));
        assert_eq!(table.rows[1].fill, None);

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn leaves_invoice_no_blank_when_match_doc_no_is_ambiguous_in_invoice_source() {
        let dir = unique_temp_path("coupons-invoice-ambiguous");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        // 两条发票明细生成相同的匹配单据号，但数电发票号码不同：属于歧义，不得任选。
        write_invoice_fixture(
            &dir,
            &[
                (
                    "销售日期:2026-08-29 单据号:收款ZFFX000003",
                    "24312000000000000001",
                ),
                (
                    "销售日期:2026-08-29 单据号:收款ZFFX000003",
                    "24312000000000000002",
                ),
            ],
        );
        write_receipts_fixture(&dir, &[]);
        write_empty_uploaded_fixtures(&dir);
        write_coupons_workbook(
            &dir,
            &[(
                "收款ZFFX000003",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "无编号",
                "10.00",
            )],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[8], Value::Empty);
        assert_eq!(
            table.rows[0].values[9],
            Value::Text("260829ZFFX000003".to_string())
        );

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn matched_remark_sinks_to_bottom_with_pink_fill_and_unmatched_rows_keep_order() {
        let dir = unique_temp_path("coupons-remark-sink");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        write_invoice_fixture(&dir, &[]);
        // “退货”销售类别直接得到非空初始备注“退货-退单”，其匹配单据号与销售明细
        // 第一行（收款ZFFX000003 / 2026-08-29）一致，构成唯一命中。
        write_receipts_fixture(&dir, &[("2026-08-29", "收款ZFFX000003", "退货", "")]);
        write_empty_uploaded_fixtures(&dir);
        write_coupons_workbook(
            &dir,
            &[
                (
                    "收款ZFFX000003",
                    "2026-08-29",
                    "商品甲",
                    "品牌甲",
                    "家电",
                    "无编号",
                    "10.00",
                ),
                (
                    "ZFFX000004",
                    "2026-08-30",
                    "商品乙",
                    "品牌乙",
                    "数码",
                    "无编号",
                    "-5.00",
                ),
            ],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(table.rows.len(), 2);

        // 未命中的第二行沉到顶部并保持原样；命中的第一行沉到底部，备注和粉色填充都到位。
        assert_eq!(
            table.rows[0].values[0],
            Value::Text("ZFFX000004".to_string())
        );
        assert_eq!(table.rows[0].values[10], Value::Text("未上传".to_string()));
        assert_eq!(table.rows[0].fill, None);

        assert_eq!(
            table.rows[1].values[0],
            Value::Text("收款ZFFX000003".to_string())
        );
        assert_eq!(
            table.rows[1].values[10],
            Value::Text("退货-退单".to_string())
        );
        assert_eq!(table.rows[1].fill, Some(Fill::Pink));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn empty_remark_in_a_hit_group_does_not_cause_ambiguity() {
        let dir = unique_temp_path("coupons-remark-blank-not-ambiguous");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        write_invoice_fixture(&dir, &[]);
        // 两条收款单统计明细命中同一个匹配单据号：一条“退货”产生非空备注“退货-退单”，
        // 一条“正常销售”且未被关联匹配命中、初始备注为空。规则 A 下空备注不参与歧义
        // 判定，去重后非空备注唯一，应正常命中“退货-退单”，不算歧义。
        write_receipts_fixture(
            &dir,
            &[
                ("2026-08-29", "收款ZFFX000003", "退货", ""),
                ("2026-08-29", "收款ZFFX000003", "正常销售", ""),
            ],
        );
        write_empty_uploaded_fixtures(&dir);
        write_coupons_workbook(
            &dir,
            &[(
                "收款ZFFX000003",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "无编号",
                "10.00",
            )],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(
            table.rows[0].values[10],
            Value::Text("退货-退单".to_string())
        );
        assert_eq!(table.rows[0].fill, Some(Fill::Pink));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn falls_through_to_not_uploaded_when_match_doc_no_has_two_distinct_nonblank_remarks() {
        let dir = unique_temp_path("coupons-remark-ambiguous");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        write_invoice_fixture(&dir, &[]);
        // 两条收款单统计明细命中同一个匹配单据号，且各自的非空初始备注不同
        // （“退货”→“退货-退单”，“零售补差”→“零售补差”）：属于歧义，不得任选，
        // 视为两阶段均未命中，备注按10.12.3节固定为“未上传”，不沉底、不上色。
        write_receipts_fixture(
            &dir,
            &[
                ("2026-08-29", "收款ZFFX000003", "退货", ""),
                ("2026-08-29", "收款ZFFX000003", "零售补差", ""),
            ],
        );
        write_empty_uploaded_fixtures(&dir);
        write_coupons_workbook(
            &dir,
            &[(
                "收款ZFFX000003",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "无编号",
                "10.00",
            )],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[10], Value::Text("未上传".to_string()));
        assert_eq!(table.rows[0].fill, None);

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn falls_back_to_uploaded_status_by_reference_when_receipts_remark_is_blank() {
        let dir = unique_temp_path("coupons-uploaded-by-reference");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        write_invoice_fixture(&dir, &[]);
        write_receipts_fixture(&dir, &[]); // 收款单统计无命中，备注留空进入兜底匹配。
        write_empty_refund_fixture(&dir, "2026年以旧换新补贴明细.xlsx");
        write_empty_refund_fixture(&dir, "2026年数码补贴明细.xlsx");
        write_uploaded_fixture(
            &dir,
            MERCHANT_APPLIANCE,
            &uploaded::APPLIANCE_TAIL,
            &[("U001", "16867252734N", "已上传", "")],
        );
        write_uploaded_fixture(&dir, MERCHANT_DIGITAL, &uploaded::DIGITAL_TAIL, &[]);
        write_coupons_workbook(
            &dir,
            &[(
                "收款ZFFX000003",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "参考号：16867252734N",
                "10.00",
            )],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[10], Value::Text("已上传".to_string()));
        assert_eq!(table.rows[0].fill, None);

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn falls_back_to_uploaded_status_by_invoice_no_when_reference_has_no_hit() {
        let dir = unique_temp_path("coupons-uploaded-by-invoice-no");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        // 摘要中没有可提取的参考号，主键（参考号）无命中，须降级到次键（数电发票号码）。
        write_invoice_fixture(
            &dir,
            &[(
                "销售日期:2026-08-29 单据号:收款ZFFX000003",
                "24312000000000000009",
            )],
        );
        write_receipts_fixture(&dir, &[]);
        write_empty_refund_fixture(&dir, "2026年以旧换新补贴明细.xlsx");
        write_empty_refund_fixture(&dir, "2026年数码补贴明细.xlsx");
        write_uploaded_fixture(&dir, MERCHANT_APPLIANCE, &uploaded::APPLIANCE_TAIL, &[]);
        write_uploaded_fixture(
            &dir,
            MERCHANT_DIGITAL,
            &uploaded::DIGITAL_TAIL,
            &[("U002", "", "已开票", "24312000000000000009")],
        );
        write_coupons_workbook(
            &dir,
            &[(
                "收款ZFFX000003",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "无编号",
                "10.00",
            )],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[10], Value::Text("已开票".to_string()));
        assert_eq!(table.rows[0].fill, None);

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reference_match_takes_priority_over_invoice_no_in_uploaded_status_lookup() {
        let dir = unique_temp_path("coupons-uploaded-priority");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        write_invoice_fixture(
            &dir,
            &[(
                "销售日期:2026-08-29 单据号:收款ZFFX000003",
                "24312000000000000009",
            )],
        );
        write_receipts_fixture(&dir, &[]);
        write_empty_refund_fixture(&dir, "2026年以旧换新补贴明细.xlsx");
        write_empty_refund_fixture(&dir, "2026年数码补贴明细.xlsx");
        // 参考号和数电发票号码分别命中不同状态：主键（参考号）命中的值必须胜出。
        write_uploaded_fixture(
            &dir,
            MERCHANT_APPLIANCE,
            &uploaded::APPLIANCE_TAIL,
            &[("U003", "16867252734N", "参考号命中", "")],
        );
        write_uploaded_fixture(
            &dir,
            MERCHANT_DIGITAL,
            &uploaded::DIGITAL_TAIL,
            &[("U004", "", "发票号命中", "24312000000000000009")],
        );
        write_coupons_workbook(
            &dir,
            &[(
                "收款ZFFX000003",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "参考号：16867252734N",
                "10.00",
            )],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(
            table.rows[0].values[10],
            Value::Text("参考号命中".to_string())
        );

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn phase_two_hit_does_not_fill_pink_and_does_not_sink() {
        let dir = unique_temp_path("coupons-phase-two-no-pink-no-sink");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        write_invoice_fixture(&dir, &[]);
        // 第一行命中第一阶段（收款单统计），将填粉色并沉底
        write_receipts_fixture(&dir, &[("2026-08-29", "收款ZFFX000001", "退货", "")]);
        write_empty_refund_fixture(&dir, "2026年以旧换新补贴明细.xlsx");
        write_empty_refund_fixture(&dir, "2026年数码补贴明细.xlsx");
        // 第二行命中第二阶段（已上传数据），不填粉色、不沉底
        write_uploaded_fixture(
            &dir,
            MERCHANT_APPLIANCE,
            &uploaded::APPLIANCE_TAIL,
            &[("U001", "16867252734N", "已上传", "")],
        );
        write_uploaded_fixture(&dir, MERCHANT_DIGITAL, &uploaded::DIGITAL_TAIL, &[]);

        write_coupons_workbook(
            &dir,
            &[
                (
                    "收款ZFFX000001",
                    "2026-08-29",
                    "商品甲",
                    "品牌甲",
                    "家电",
                    "无编号",
                    "10.00",
                ),
                (
                    "收款ZFFX000002",
                    "2026-08-30",
                    "商品乙",
                    "品牌乙",
                    "家电",
                    "参考号：16867252734N",
                    "20.00",
                ),
                (
                    "收款ZFFX000003",
                    "2026-08-31",
                    "商品丙",
                    "品牌丙",
                    "家电",
                    "无编号",
                    "30.00",
                ),
            ],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(table.rows.len(), 3);

        // 前部：第二行（第二阶段命中）与第三行（未命中），保持在顶部且无填色
        assert_eq!(
            table.rows[0].values[0],
            Value::Text("收款ZFFX000002".to_string())
        );
        assert_eq!(table.rows[0].values[10], Value::Text("已上传".to_string()));
        assert_eq!(table.rows[0].fill, None);

        assert_eq!(
            table.rows[1].values[0],
            Value::Text("收款ZFFX000003".to_string())
        );
        assert_eq!(table.rows[1].values[10], Value::Text("未上传".to_string()));
        assert_eq!(table.rows[1].fill, None);

        // 底部：第一行（第一阶段命中），填粉色并沉底
        assert_eq!(
            table.rows[2].values[0],
            Value::Text("收款ZFFX000001".to_string())
        );
        assert_eq!(
            table.rows[2].values[10],
            Value::Text("退货-退单".to_string())
        );
        assert_eq!(table.rows[2].fill, Some(Fill::Pink));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn falls_through_to_not_uploaded_when_uploaded_reference_match_is_ambiguous() {
        let dir = unique_temp_path("coupons-uploaded-ambiguous");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        // 数电发票号码同样能命中，但主键歧义须立即判定未命中，不得降级到次键。
        write_invoice_fixture(
            &dir,
            &[(
                "销售日期:2026-08-29 单据号:收款ZFFX000003",
                "24312000000000000009",
            )],
        );
        write_receipts_fixture(&dir, &[]);
        write_empty_refund_fixture(&dir, "2026年以旧换新补贴明细.xlsx");
        write_empty_refund_fixture(&dir, "2026年数码补贴明细.xlsx");
        write_uploaded_fixture(
            &dir,
            MERCHANT_APPLIANCE,
            &uploaded::APPLIANCE_TAIL,
            &[
                ("U005", "16867252734N", "状态甲", ""),
                ("U006", "16867252734N", "状态乙", ""),
            ],
        );
        write_uploaded_fixture(
            &dir,
            MERCHANT_DIGITAL,
            &uploaded::DIGITAL_TAIL,
            &[("U007", "", "发票号命中", "24312000000000000009")],
        );
        write_coupons_workbook(
            &dir,
            &[(
                "收款ZFFX000003",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "参考号：16867252734N",
                "10.00",
            )],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[10], Value::Text("未上传".to_string()));
        assert_eq!(table.rows[0].fill, None);

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_uploaded_source_missing() {
        let dir = unique_temp_path("coupons-missing-uploaded");
        std::fs::create_dir_all(&dir).unwrap();
        write_unionpay_fixture(&dir, "16867252734N");
        write_invoice_fixture(&dir, &[]);
        write_receipts_fixture(&dir, &[]);
        write_coupons_workbook(
            &dir,
            &[(
                "Z1",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "无编号",
                "10.00",
            )],
        );

        // 已上传家电电脑/已上传数码样本文件缺失：必须停止，不得让备注兜底匹配静默跳过。
        let error = CouponsJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_title_mismatch() {
        let dir = unique_temp_path("coupons-bad-title");
        std::fs::create_dir_all(&dir).unwrap();
        write_unionpay_fixture(&dir, "16867252734N");
        write_invoice_fixture(&dir, &[]);
        write_receipts_fixture(&dir, &[]);
        write_empty_uploaded_fixtures(&dir);

        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, "错误标题").unwrap();
        for (col, header) in SOURCE_HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        sheet.write_string(2, 0, "合计").unwrap();
        workbook.save(dir.join(FILE_NAME)).unwrap();

        let error = CouponsJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Structure { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_own_structure_error_before_checking_external_dependencies() {
        let dir = unique_temp_path("coupons-own-file-checked-first");
        std::fs::create_dir_all(&dir).unwrap();
        // 输入目录完全空白：销售用券情况统计.xlsx 与全部外部依赖（门店银联、发票明细、
        // 收款单统计、已上传数据）均缺失。按10.1节顺序，须先报告本任务自身输入文件缺失，
        // 不得让外部依赖缺失的错误抢先出现，掩盖了真正的问题（本文件根本不存在）。
        let error = CouponsJob.run(&dir).unwrap_err();
        match error {
            ProcessError::NoInput { pattern } => assert_eq!(pattern, FILE_NAME),
            other => panic!("expected NoInput for {FILE_NAME}, got {other:?}"),
        }

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_authority_source_missing() {
        let dir = unique_temp_path("coupons-missing-authority");
        std::fs::create_dir_all(&dir).unwrap();
        write_coupons_workbook(
            &dir,
            &[(
                "Z1",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "无编号",
                "10.00",
            )],
        );

        // 门店银联样本文件缺失：必须停止，不得绕过权威校验集。
        let error = CouponsJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_invoice_source_missing() {
        let dir = unique_temp_path("coupons-missing-invoice");
        std::fs::create_dir_all(&dir).unwrap();
        write_unionpay_fixture(&dir, "16867252734N");
        write_coupons_workbook(
            &dir,
            &[(
                "Z1",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "无编号",
                "10.00",
            )],
        );

        // 发票明细样本文件缺失：必须停止，不得让数电发票号码列全部留空后继续输出。
        let error = CouponsJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_receipts_source_missing() {
        let dir = unique_temp_path("coupons-missing-receipts");
        std::fs::create_dir_all(&dir).unwrap();
        write_unionpay_fixture(&dir, "16867252734N");
        write_invoice_fixture(&dir, &[]);
        write_coupons_workbook(
            &dir,
            &[(
                "Z1",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "无编号",
                "10.00",
            )],
        );

        // 收款单统计样本文件缺失：必须停止，不得让备注列全部留空后继续输出。
        let error = CouponsJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_empty_subsidy_amount() {
        let dir = unique_temp_path("coupons-empty-subsidy");
        std::fs::create_dir_all(&dir).unwrap();
        write_unionpay_fixture(&dir, "16867252734N");
        write_invoice_fixture(&dir, &[]);
        write_receipts_fixture(&dir, &[]);
        write_empty_uploaded_fixtures(&dir);

        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, TITLE).unwrap();
        for (col, header) in SOURCE_HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        sheet
            .write_string(2, (COL_DOC_NO - 1) as u16, "Z1")
            .unwrap();
        // 合计（补贴额）留空。
        sheet.write_string(3, 0, "合计").unwrap();
        workbook.save(dir.join(FILE_NAME)).unwrap();

        let error = CouponsJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Data { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }
}
