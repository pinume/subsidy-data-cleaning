use std::collections::{HashMap, HashSet};
use std::path::{Path, PathBuf};

use rust_decimal::Decimal;
use rust_decimal::prelude::FromPrimitive;

use crate::io::paths::list_xlsx_files;
use crate::io::xlsx_reader::{RawCell, SheetGrid, open_sheets};
use crate::model::{Column, ColumnType, DecimalScale, Fill, ProcessError, Row, Table, Value};

use super::{
    Category, Job, amount_value, cell_amount, cell_text, data_error, pick_unique_latest,
    resolve_synonym_column, text_value,
};

pub(crate) const OUTPUT_FIELDS: [&str; 24] = [
    "拨付批次",
    "交易完成时间",
    "交易参考号",
    "商户订单号",
    "交易订单号",
    "销售企业名称",
    "核销商编",
    "其他支付",
    "销售金额",
    "实收销售金额",
    "补贴金额",
    "补贴比例",
    "SN码",
    "所在地区",
    "商品编码",
    "能耗等级",
    "编码品类",
    "商品名称",
    "发票金额",
    "发票号",
    "发票头/购买方名称",
    "ID",
    "退回原因",
    "原拨付批次",
];

const COLUMN_TYPES: [ColumnType; 24] = [
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Ratio,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
];

// 24 列固定顺序中的关键索引（0 基）。
const OTHER_PAYMENT: usize = 7;
const SUBSIDY_AMOUNT: usize = 10;
const RATIO: usize = 11;
const DEALER_CODE: usize = 6; // 核销商编

struct RefundConfig {
    category: Category,
    title: &'static str,
    output_stem: &'static str,
    /// 文件名后缀，如`年以旧换新补贴明细.xlsx`；前面须为 4 位数字年份。
    filename_suffix: &'static str,
    /// 每个统一字段（按 24 列顺序）对应的已确认同义源字段名候选集合。
    field_synonyms: [&'static [&'static str]; 24],
    /// 判定"批次明细表"的基础字段索引；缺失任一项即为待映射异常，终止处理。
    required_indices: &'static [usize],
    /// 重复分组依据字段的索引，按优先级排列。
    grouping_priority: [usize; 3],
    /// 只保留`核销商编`精确等于该值的明细行，其余门店编码的记录在读取阶段即排除，
    /// 不参与后续合并、去重或输出。
    dealer_code: &'static str,
}

const APPLIANCE_SYNONYMS: [&[&str]; 24] = [
    &["拨付批次"],
    &["交易完成时间"],
    &["交易参考号"],
    &["商户订单号"],
    &["交易订单号"],
    &["销方名称", "销售企业名称"],
    &["商户编号", "核销商编"],
    &["其他支付"],
    &["应收销售金额", "销售金额"],
    &["实收销售金额"],
    &["补贴金额"],
    &["补贴比例"],
    &["SN码"],
    &["所在地区"],
    &["商品编码"],
    &["能耗等级"],
    &["编码品类"],
    &["商品名称"],
    &["发票金额"],
    &["发票号"],
    &["发票头/购买方名称"],
    &["ID"],
    &["退回原因"],
    &["原拨付批次"],
];

const APPLIANCE_REQUIRED: [usize; 19] = [
    0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 19, 21,
];

const APPLIANCE_CONFIG: RefundConfig = RefundConfig {
    category: Category::RefundAppliance,
    title: "家电电脑回款明细",
    output_stem: "回款明细家电电脑",
    filename_suffix: "年以旧换新补贴明细.xlsx",
    field_synonyms: APPLIANCE_SYNONYMS,
    required_indices: &APPLIANCE_REQUIRED,
    grouping_priority: [4, 3, 19], // 交易订单号 → 商户订单号 → 发票号
    dealer_code: "89813015722APT1",
};

const DIGITAL_SYNONYMS: [&[&str]; 24] = [
    &["拨付批次"],
    &["交易完成时间"],
    &["参考号", "交易参考号"],
    &["商户订单号"],
    &["订单号", "交易订单号"],
    &["销方名称", "销售企业名称"],
    &["经销商编号", "商户编号", "核销商编"],
    &["其他支付"],
    &["应收销售金额", "销售金额"],
    &["实收销售金额"],
    &["补贴销售金额", "补贴金额"],
    &["补贴比例/补贴限额", "补贴比例"],
    &["SN码"],
    &["所在地区"],
    &["商品编码"],
    &["能耗等级"],
    &["品类", "类别", "编码品类"],
    &["商品明细", "商品名称"],
    &["发票金额"],
    &["发票号码", "发票号"],
    &["发票头/购买方名称"],
    &["ID"],
    &["退回原因"],
    &["原拨付批次"],
];

const DIGITAL_REQUIRED: [usize; 13] = [0, 2, 3, 5, 6, 8, 10, 12, 13, 14, 16, 17, 19];

const DIGITAL_CONFIG: RefundConfig = RefundConfig {
    category: Category::RefundDigital,
    title: "数码回款明细",
    output_stem: "回款明细数码",
    filename_suffix: "年数码补贴明细.xlsx",
    field_synonyms: DIGITAL_SYNONYMS,
    required_indices: &DIGITAL_REQUIRED,
    grouping_priority: [2, 3, 19], // 交易参考号 → 商户订单号 → 发票号
    dealer_code: "89813014812B06R",
};

pub struct RefundJob(&'static RefundConfig);

pub const REFUND_APPLIANCE: RefundJob = RefundJob(&APPLIANCE_CONFIG);
pub const REFUND_DIGITAL: RefundJob = RefundJob(&DIGITAL_CONFIG);

impl Job for RefundJob {
    fn category(&self) -> Category {
        self.0.category
    }

    fn title(&self) -> &'static str {
        self.0.title
    }

    fn output_stem(&self) -> &'static str {
        self.0.output_stem
    }

    fn run(&self, input_dir: &Path) -> Result<Table, ProcessError> {
        run_refund(self.0, input_dir)
    }
}

/// 文件名须为`<4 位年份><后缀>`，如`2026年以旧换新补贴明细.xlsx`。
fn matches_filename(name: &str, suffix: &str) -> bool {
    match name.strip_suffix(suffix) {
        Some(year) => year.len() == 4 && year.bytes().all(|b| b.is_ascii_digit()),
        None => false,
    }
}

/// 比例字段：接受十进制数值，或带`%`的文本（按 100 换算为底层比例）。
fn cell_ratio(cell: &RawCell) -> Result<Option<Decimal>, String> {
    match cell {
        RawCell::Empty => Ok(None),
        RawCell::Text(t) if t.trim().is_empty() => Ok(None),
        RawCell::Text(t) => {
            let trimmed = t.trim();
            let (numeric, is_percent) = match trimmed.strip_suffix('%') {
                Some(rest) => (rest.trim(), true),
                None => (trimmed, false),
            };
            let value: Decimal = numeric
                .parse()
                .map_err(|_| format!("文本“{t}”无法解析为比例"))?;
            Ok(Some(if is_percent {
                value / Decimal::from(100)
            } else {
                value
            }))
        }
        RawCell::Int(n) => Ok(Some(Decimal::from(*n))),
        RawCell::Float(f) => Decimal::from_f64(*f)
            .map(Some)
            .ok_or_else(|| format!("数值 {f} 无法转换为比例")),
        other => Err(format!("比例字段出现非数值内容：{other}")),
    }
}

/// 其他支付：把文本`-`原样保留为`Value::Text("-")`，否则按普通数值字段处理。
fn read_other_payment(cell: &RawCell) -> Result<Value, String> {
    if let RawCell::Text(t) = cell
        && t.trim() == "-"
    {
        return Ok(Value::Text("-".to_string()));
    }
    cell_amount(cell).map(amount_value)
}

/// 文本字段：与共享的`cell_text`相比，遇到 Excel 错误值（如`#N/A`）时原样保留为文本，
/// 不终止处理——已确认：源表文本字段可能出现查找失败等错误值，按原样输出，不做推断。
/// `核销商编`已在读取阶段按商户号过滤为匹配值，不会经此路径遇到错误值。
fn cell_text_tolerant(cell: &RawCell) -> Result<String, String> {
    match cell {
        RawCell::Error(_) => Ok(cell.to_string()),
        other => cell_text(other),
    }
}

fn read_row(
    sheet: &SheetGrid,
    row: u32,
    columns: &[Option<u32>],
    file: &str,
    sheet_name: &str,
) -> Result<Vec<Value>, ProcessError> {
    let cell_at = |index: usize| columns[index].map_or(RawCell::Empty, |col| sheet.cell(row, col));

    let mut values = Vec::with_capacity(24);
    for index in 0..24 {
        let cell = cell_at(index);
        let field = OUTPUT_FIELDS[index];
        let value = if index == OTHER_PAYMENT {
            read_other_payment(&cell).map_err(|detail| {
                data_error(file, sheet_name, row, field, cell.to_string(), detail)
            })?
        } else if index == SUBSIDY_AMOUNT {
            let amount = cell_amount(&cell)
                .map_err(|detail| {
                    data_error(file, sheet_name, row, field, cell.to_string(), detail)
                })?
                .ok_or_else(|| {
                    data_error(
                        file,
                        sheet_name,
                        row,
                        field,
                        cell.to_string(),
                        "补贴金额为空或无法解析，不得参与求和".to_string(),
                    )
                })?;
            Value::Decimal(amount)
        } else if index == RATIO {
            cell_ratio(&cell)
                .map(|ratio| ratio.map_or(Value::Empty, Value::Ratio))
                .map_err(|detail| {
                    data_error(file, sheet_name, row, field, cell.to_string(), detail)
                })?
        } else if matches!(COLUMN_TYPES[index], ColumnType::Decimal(_)) {
            cell_amount(&cell).map(amount_value).map_err(|detail| {
                data_error(file, sheet_name, row, field, cell.to_string(), detail)
            })?
        } else {
            text_value(cell_text_tolerant(&cell).map_err(|detail| {
                data_error(file, sheet_name, row, field, cell.to_string(), detail)
            })?)
        };
        values.push(value);
    }
    Ok(values)
}

/// 分组依据：优先级中首个非空文本字段的`(字段索引, 去首尾空白后的值)`；
/// 索引本身即代表"字段类型"，天然避免不同字段交叉匹配。
fn grouping_key(values: &[Value], priority: [usize; 3]) -> Option<(usize, String)> {
    for index in priority {
        if let Value::Text(text) = &values[index] {
            let trimmed = text.trim();
            if !trimmed.is_empty() {
                return Some((index, trimmed.to_string()));
            }
        }
    }
    None
}

fn subsidy_amount(values: &[Value]) -> Decimal {
    match &values[SUBSIDY_AMOUNT] {
        Value::Decimal(amount) => *amount,
        _ => unreachable!("补贴金额在读取阶段已确保为 Decimal"),
    }
}

/// 第 7.5/8.6 节：跨批次统一分组、按补贴金额合计沉底；未沉底记录在前，沉底记录在后，
/// 两个区域内部均保持合并后的原相对顺序。
fn classify_and_sink(records: Vec<Vec<Value>>, priority: [usize; 3]) -> Vec<Row> {
    let mut groups: HashMap<(usize, String), Vec<usize>> = HashMap::new();
    for (index, values) in records.iter().enumerate() {
        if let Some(key) = grouping_key(values, priority) {
            groups.entry(key).or_default().push(index);
        }
    }

    let mut duplicate_members: HashSet<usize> = HashSet::new();
    let mut sink: HashSet<usize> = HashSet::new();
    for indices in groups.values() {
        if indices.len() < 2 {
            continue;
        }
        duplicate_members.extend(indices.iter().copied());
        let total = indices
            .iter()
            .fold(Decimal::ZERO, |acc, &i| acc + subsidy_amount(&records[i]));
        if total <= Decimal::ZERO {
            sink.extend(indices.iter().copied());
        } else {
            let last = *indices.iter().max().unwrap();
            sink.extend(indices.iter().copied().filter(|&i| i != last));
        }
    }
    for (index, values) in records.iter().enumerate() {
        if !duplicate_members.contains(&index) && subsidy_amount(values) < Decimal::ZERO {
            sink.insert(index);
        }
    }

    let mut normal = Vec::new();
    let mut sunk = Vec::new();
    for (index, values) in records.into_iter().enumerate() {
        if sink.contains(&index) {
            sunk.push(Row {
                values,
                fill: Some(Fill::Pink),
            });
        } else {
            normal.push(Row { values, fill: None });
        }
    }
    normal.into_iter().chain(sunk).collect()
}

fn output_columns() -> Vec<Column> {
    OUTPUT_FIELDS
        .iter()
        .zip(COLUMN_TYPES)
        .map(|(&name, ty)| Column { name, ty })
        .collect()
}

/// 在候选文件中选择文件名年份最新的一个；不合并较早年份的文件。
fn select_latest_file(candidates: Vec<PathBuf>, suffix: &str) -> Result<PathBuf, ProcessError> {
    let dated: Vec<(PathBuf, u32)> = candidates
        .into_iter()
        .map(|path| {
            let year = path
                .file_name()
                .and_then(|n| n.to_str())
                .and_then(|n| n.strip_suffix(suffix))
                .and_then(|y| y.parse::<u32>().ok())
                .expect("candidates are pre-filtered by matches_filename");
            (path, year)
        })
        .collect();

    pick_unique_latest(dated, "最新年份无法唯一确定：多个文件对应同一最新年份")
}

fn run_refund(config: &RefundConfig, input_dir: &Path) -> Result<Table, ProcessError> {
    let candidates: Vec<PathBuf> = list_xlsx_files(input_dir)?
        .into_iter()
        .filter(|path| {
            path.file_name()
                .and_then(|n| n.to_str())
                .is_some_and(|n| matches_filename(n, config.filename_suffix))
        })
        .collect();

    if candidates.is_empty() {
        return Err(ProcessError::NoInput {
            pattern: format!("yyyy{}", config.filename_suffix),
        });
    }
    let path = select_latest_file(candidates, config.filename_suffix)?;
    let file_name = path.file_name().unwrap().to_string_lossy().into_owned();

    let sheets = open_sheets(&path)?;
    let mut records: Vec<Vec<Value>> = Vec::new();

    for sheet in &sheets {
        if sheet.name() == "汇总" {
            continue;
        }
        let sheet_name = sheet.name().to_string();
        let header = sheet.row_texts(1);

        let mut columns: Vec<Option<u32>> = Vec::with_capacity(24);
        for synonyms in &config.field_synonyms {
            let resolved = resolve_synonym_column(&header, synonyms).map_err(|detail| {
                ProcessError::Structure {
                    file: file_name.clone(),
                    sheet: sheet_name.clone(),
                    detail,
                }
            })?;
            columns.push(resolved);
        }

        let missing: Vec<&str> = config
            .required_indices
            .iter()
            .filter(|&&i| columns[i].is_none())
            .map(|&i| OUTPUT_FIELDS[i])
            .collect();
        if !missing.is_empty() {
            return Err(ProcessError::Structure {
                file: file_name,
                sheet: sheet_name,
                detail: format!("待映射异常：缺少基础字段 {}", missing.join("、")),
            });
        }

        let dealer_col = columns[DEALER_CODE].expect("核销商编在 required_indices 中");
        let last_row = sheet.last_value_row().unwrap_or(1);
        for row in 2..=last_row {
            if sheet.cell(row, dealer_col).to_string().trim() != config.dealer_code {
                continue;
            }
            records.push(read_row(sheet, row, &columns, &file_name, &sheet_name)?);
        }
    }

    let rows = classify_and_sink(records, config.grouping_priority);
    Ok(Table {
        columns: output_columns(),
        rows,
    })
}

#[cfg(test)]
mod tests {
    use rust_xlsxwriter::{Formula, Workbook};

    use super::*;
    use crate::test_support::unique_temp_path;

    type SheetSpec<'a> = (&'a str, &'a [&'a str], &'a [Vec<String>]);

    fn write_workbook(path: &Path, sheets: &[SheetSpec]) {
        let mut workbook = Workbook::new();
        for (name, header, rows) in sheets {
            let sheet = workbook.add_worksheet();
            sheet.set_name(*name).unwrap();
            for (col, value) in header.iter().enumerate() {
                sheet.write_string(0, col as u16, *value).unwrap();
            }
            for (row_index, row) in rows.iter().enumerate() {
                for (col, value) in row.iter().enumerate() {
                    sheet
                        .write_string((1 + row_index) as u32, col as u16, value.as_str())
                        .unwrap();
                }
            }
        }
        workbook.save(path).unwrap();
    }

    const APPLIANCE_HEADER: [&str; 19] = [
        "拨付批次",
        "交易参考号",
        "商户订单号",
        "交易订单号",
        "销售企业名称",
        "核销商编",
        "其他支付",
        "销售金额",
        "实收销售金额",
        "补贴金额",
        "补贴比例",
        "SN码",
        "所在地区",
        "商品编码",
        "能耗等级",
        "编码品类",
        "商品名称",
        "发票号",
        "ID",
    ];

    /// 按`APPLIANCE_HEADER`顺序生成一行；`label`写入拨付批次，便于测试用它识别行。
    fn appliance_row<'a>(
        label: &'a str,
        merchant_order_no: &'a str,
        transaction_order_no: &'a str,
        subsidy: &'a str,
    ) -> Vec<String> {
        vec![
            label.to_string(),
            String::new(),
            merchant_order_no.to_string(),
            transaction_order_no.to_string(),
            "企业甲".to_string(),
            APPLIANCE_CONFIG.dealer_code.to_string(),
            "0.00".to_string(),
            "100.00".to_string(),
            "90.00".to_string(),
            subsidy.to_string(),
            "0.15".to_string(),
            "SN001".to_string(),
            "地区甲".to_string(),
            "CODE001".to_string(),
            "一级".to_string(),
            "品类甲".to_string(),
            "商品甲".to_string(),
            "INV001".to_string(),
            "ID001".to_string(),
        ]
    }

    fn label_of(row: &Row) -> &str {
        match &row.values[0] {
            Value::Text(text) => text.as_str(),
            _ => panic!("expected text"),
        }
    }

    #[test]
    fn matches_filename_requires_four_digit_year_prefix() {
        assert!(matches_filename(
            "2026年以旧换新补贴明细.xlsx",
            "年以旧换新补贴明细.xlsx"
        ));
        assert!(!matches_filename(
            "26年以旧换新补贴明细.xlsx",
            "年以旧换新补贴明细.xlsx"
        ));
        assert!(!matches_filename(
            "2026年数码补贴明细.xlsx",
            "年以旧换新补贴明细.xlsx"
        ));
    }

    #[test]
    fn merges_batches_with_synonym_headers_and_excludes_summary_sheet() {
        let dir = unique_temp_path("refund-appliance-happy-path");
        std::fs::create_dir_all(&dir).unwrap();

        // 第一批次用“销方名称”，第二批次用“销售企业名称”——同义字段必须都能映射。
        let mut header_batch1 = APPLIANCE_HEADER;
        header_batch1[4] = "销方名称";

        write_workbook(
            &dir.join("2026年以旧换新补贴明细.xlsx"),
            &[
                (
                    "批次一",
                    &header_batch1,
                    &[appliance_row("R1", "M1", "T1", "10.00")],
                ),
                (
                    "批次二",
                    &APPLIANCE_HEADER,
                    &[appliance_row("R2", "M2", "T2", "20.00")],
                ),
                ("汇总", &["其他列"], &[vec!["不应被读取".to_string()]]),
            ],
        );

        let table = REFUND_APPLIANCE.run(&dir).unwrap();

        assert_eq!(table.columns.len(), 24);
        assert_eq!(table.rows.len(), 2);
        assert_eq!(label_of(&table.rows[0]), "R1");
        assert_eq!(label_of(&table.rows[1]), "R2");
        // 退回原因、原拨付批次源表无对应字段，留空。
        assert_eq!(table.rows[0].values[22], Value::Empty);
        assert_eq!(table.rows[0].values[23], Value::Empty);
        // 销方名称已统一映射为销售企业名称的值。
        assert_eq!(table.rows[0].values[5], Value::Text("企业甲".to_string()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn maps_new_optional_fields_when_present_and_blank_when_absent() {
        let dir = unique_temp_path("refund-optional-new-fields");
        std::fs::create_dir_all(&dir).unwrap();

        let mut header_with_new_fields = APPLIANCE_HEADER.to_vec();
        header_with_new_fields.insert(1, "交易完成时间");
        header_with_new_fields.insert(18, "发票金额"); // 插入到“发票号”之前
        header_with_new_fields.insert(20, "发票头/购买方名称"); // 插入到“ID”之前

        let mut row_with_new_fields = appliance_row("WITH", "M1", "T1", "10.00");
        row_with_new_fields.insert(1, "2026-01-01 20:14:22".to_string());
        row_with_new_fields.insert(18, "6129".to_string());
        row_with_new_fields.insert(20, "王惠霞".to_string());

        write_workbook(
            &dir.join("2026年以旧换新补贴明细.xlsx"),
            &[
                ("有新字段", &header_with_new_fields, &[row_with_new_fields]),
                (
                    "无新字段",
                    &APPLIANCE_HEADER,
                    &[appliance_row("WITHOUT", "M2", "T2", "10.00")],
                ),
            ],
        );

        let table = REFUND_APPLIANCE.run(&dir).unwrap();
        let with_row = table.rows.iter().find(|r| label_of(r) == "WITH").unwrap();
        assert_eq!(
            with_row.values[1],
            Value::Text("2026-01-01 20:14:22".to_string())
        );
        assert_eq!(with_row.values[18], Value::Decimal("6129".parse().unwrap()));
        assert_eq!(with_row.values[20], Value::Text("王惠霞".to_string()));

        let without_row = table
            .rows
            .iter()
            .find(|r| label_of(r) == "WITHOUT")
            .unwrap();
        assert_eq!(without_row.values[1], Value::Empty);
        assert_eq!(without_row.values[18], Value::Empty);
        assert_eq!(without_row.values[20], Value::Empty);

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn drops_rows_whose_dealer_code_does_not_match_configured_merchant() {
        let dir = unique_temp_path("refund-dealer-code-filter");
        std::fs::create_dir_all(&dir).unwrap();

        let mut other_dealer_row = appliance_row("OTHER", "M2", "T2", "20.00");
        other_dealer_row[5] = "89813015722APOTHER".to_string();

        write_workbook(
            &dir.join("2026年以旧换新补贴明细.xlsx"),
            &[(
                "批次一",
                &APPLIANCE_HEADER,
                &[appliance_row("KEEP", "M1", "T1", "10.00"), other_dealer_row],
            )],
        );

        let table = REFUND_APPLIANCE.run(&dir).unwrap();

        assert_eq!(table.rows.len(), 1);
        assert_eq!(label_of(&table.rows[0]), "KEEP");

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn excel_error_value_in_text_field_is_preserved_instead_of_erroring() {
        let dir = unique_temp_path("refund-na-preserved");
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("2026年以旧换新补贴明细.xlsx");

        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("批次一").unwrap();
        for (col, header) in APPLIANCE_HEADER.iter().enumerate() {
            sheet.write_string(0, col as u16, *header).unwrap();
        }
        let row = appliance_row("R1", "M1", "T1", "10.00");
        for (col, value) in row.iter().enumerate() {
            sheet.write_string(1, col as u16, value.as_str()).unwrap();
        }
        // 所在地区（第 13 列，0 基列号 12）：真实数据中曾在文本字段出现过的查找失败错误值；
        // 核销商编本身已由商户号过滤保证为匹配值，不会是错误值，故用其他文本字段验证容错。
        sheet
            .write_formula(1, 12, Formula::new("=NA()").set_result("#N/A"))
            .unwrap();
        workbook.save(&path).unwrap();

        let table = REFUND_APPLIANCE.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[13], Value::Text("#N/A".to_string()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_missing_required_field() {
        let dir = unique_temp_path("refund-missing-field");
        std::fs::create_dir_all(&dir).unwrap();
        let mut header = APPLIANCE_HEADER.to_vec();
        header.pop(); // 去掉 ID
        write_workbook(
            &dir.join("2026年以旧换新补贴明细.xlsx"),
            &[(
                "批次一",
                &header,
                &[appliance_row("R1", "M1", "T1", "10.00")
                    .into_iter()
                    .take(header.len())
                    .collect()],
            )],
        );

        let error = REFUND_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Structure { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_empty_subsidy_amount() {
        let dir = unique_temp_path("refund-empty-subsidy");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("2026年以旧换新补贴明细.xlsx"),
            &[(
                "批次一",
                &APPLIANCE_HEADER,
                &[appliance_row("R1", "M1", "T1", "")],
            )],
        );

        let error = REFUND_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Data { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn groups_across_sheets_and_sinks_standalone_negative_records() {
        let dir = unique_temp_path("refund-grouping");
        std::fs::create_dir_all(&dir).unwrap();

        write_workbook(
            &dir.join("2026年以旧换新补贴明细.xlsx"),
            &[
                (
                    "批次一",
                    &APPLIANCE_HEADER,
                    &[
                        appliance_row("A", "", "DUP1", "10.00"),
                        appliance_row("B", "", "DUP1", "-20.00"), // 组内合计 -10 <= 0，全部沉底
                        appliance_row("C", "", "DUP2", "30.00"),
                    ],
                ),
                (
                    "批次二",
                    &APPLIANCE_HEADER,
                    &[
                        appliance_row("D", "", "DUP2", "40.00"), // 跨批次同组，合计 70 > 0，保留最后一次（D）
                        appliance_row("E", "UNIQUE1", "", "5.00"), // 单条记录不受影响
                        appliance_row("F", "", "CROSS", "1.00"), // 交易订单号=CROSS
                        appliance_row("G", "CROSS", "", "1.00"), // 商户订单号=CROSS，字段类型不同，不得与 F 同组
                        appliance_row("H", "UNIQUE2", "", "-5.00"), // 未形成重复组的负数记录沉底
                    ],
                ),
            ],
        );

        let table = REFUND_APPLIANCE.run(&dir).unwrap();
        assert_eq!(table.rows.len(), 8);

        let labels: Vec<&str> = table.rows.iter().map(label_of).collect();
        assert_eq!(labels, vec!["D", "E", "F", "G", "A", "B", "C", "H"]);

        for label in ["D", "E", "F", "G"] {
            let row = table.rows.iter().find(|r| label_of(r) == label).unwrap();
            assert_eq!(row.fill, None, "{label} 不应沉底");
        }
        for label in ["A", "B", "C", "H"] {
            let row = table.rows.iter().find(|r| label_of(r) == label).unwrap();
            assert_eq!(row.fill, Some(Fill::Pink), "{label} 应沉底");
        }

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn parses_ratio_from_percent_text() {
        let dir = unique_temp_path("refund-ratio-percent");
        std::fs::create_dir_all(&dir).unwrap();
        let mut row = appliance_row("R1", "M1", "T1", "10.00");
        row[10] = "15%".to_string();
        write_workbook(
            &dir.join("2026年以旧换新补贴明细.xlsx"),
            &[("批次一", &APPLIANCE_HEADER, &[row])],
        );

        let table = REFUND_APPLIANCE.run(&dir).unwrap();
        match table.rows[0].values[RATIO] {
            Value::Ratio(ratio) => assert_eq!(ratio, "0.15".parse().unwrap()),
            _ => panic!("expected ratio"),
        }

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn other_payment_dash_is_allowed_for_both_categories() {
        let appliance_dir = unique_temp_path("refund-dash-appliance");
        std::fs::create_dir_all(&appliance_dir).unwrap();
        let mut row = appliance_row("R1", "M1", "T1", "10.00");
        row[6] = "-".to_string();
        write_workbook(
            &appliance_dir.join("2026年以旧换新补贴明细.xlsx"),
            &[("批次一", &APPLIANCE_HEADER, &[row])],
        );
        let table = REFUND_APPLIANCE.run(&appliance_dir).unwrap();
        assert_eq!(
            table.rows[0].values[OTHER_PAYMENT],
            Value::Text("-".to_string())
        );
        std::fs::remove_dir_all(&appliance_dir).unwrap();

        let digital_header: [&str; 14] = [
            "拨付批次",
            "参考号",
            "商户订单号",
            "销售企业名称",
            "核销商编",
            "其他支付",
            "销售金额",
            "补贴金额",
            "SN码",
            "所在地区",
            "商品编码",
            "编码品类",
            "商品名称",
            "发票号",
        ];
        let digital_row = vec![
            "R1".to_string(),
            "REF1".to_string(),
            "M1".to_string(),
            "企业甲".to_string(),
            DIGITAL_CONFIG.dealer_code.to_string(),
            "-".to_string(),
            "100.00".to_string(),
            "10.00".to_string(),
            "SN001".to_string(),
            "地区甲".to_string(),
            "CODE001".to_string(),
            "品类甲".to_string(),
            "商品甲".to_string(),
            "INV001".to_string(),
        ];
        let digital_dir = unique_temp_path("refund-dash-digital");
        std::fs::create_dir_all(&digital_dir).unwrap();
        write_workbook(
            &digital_dir.join("2026年数码补贴明细.xlsx"),
            &[("批次一", &digital_header, &[digital_row])],
        );
        let table = REFUND_DIGITAL.run(&digital_dir).unwrap();
        assert_eq!(
            table.rows[0].values[OTHER_PAYMENT],
            Value::Text("-".to_string())
        );
        std::fs::remove_dir_all(&digital_dir).unwrap();
    }

    #[test]
    fn selects_latest_year_and_ignores_earlier_files() {
        let dir = unique_temp_path("refund-multiple-years");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("2025年以旧换新补贴明细.xlsx"),
            &[(
                "批次一",
                &APPLIANCE_HEADER,
                &[appliance_row("OLD", "M1", "T1", "10.00")],
            )],
        );
        write_workbook(
            &dir.join("2026年以旧换新补贴明细.xlsx"),
            &[(
                "批次一",
                &APPLIANCE_HEADER,
                &[appliance_row("NEW", "M2", "T2", "10.00")],
            )],
        );

        let table = REFUND_APPLIANCE.run(&dir).unwrap();
        assert_eq!(table.rows.len(), 1);
        assert_eq!(label_of(&table.rows[0]), "NEW");

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_nothing_matches() {
        let dir = unique_temp_path("refund-no-input");
        std::fs::create_dir_all(&dir).unwrap();

        let error = REFUND_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }
}
