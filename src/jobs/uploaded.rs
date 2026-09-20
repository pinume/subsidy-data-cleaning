use std::collections::HashMap;
use std::collections::HashSet;
use std::path::{Path, PathBuf};
use std::str::FromStr;

use rust_decimal::Decimal;

use crate::io::paths::list_xlsx_files;
use crate::io::xlsx_reader::{RawCell, SheetGrid, open_sheets};
use crate::model::{Column, ColumnType, DecimalScale, Fill, ProcessError, Row, Table, Value};

use super::refund::{REFUND_APPLIANCE, REFUND_DIGITAL};
use super::{
    Category, Job, MultiValueIndex, PriorityOutcome, amount_value, cell_amount, cell_date_or_text,
    cell_datetime_or_text, cell_text, check_duplicate_fingerprint, data_error,
    resolve_synonym_column, resolve_via, text_value,
};

/// 前25列两组数据组结构完全相同，按固定列位置读取。
const FRONT_LEN: usize = 25;
const FIELD_COUNT: usize = 58;
const TAIL_LEN: usize = FIELD_COUNT - FRONT_LEN;

pub(crate) const FRONT_HEADERS: [&str; FRONT_LEN] = [
    "实时清分UUID",
    "商户号",
    "商户名称",
    "订单号",
    "交易日期",
    "交易金额",
    "检索参考号",
    "模版类型",
    "状态",
    "描述",
    "提交时间",
    "更新时间",
    "终端号",
    "分店id",
    "分店名",
    "所在地区",
    "详细地址",
    "地区编码",
    "tel",
    "发票号码",
    "发票金额",
    "购买方名称",
    "图片1",
    "S/N码",
    "是否属于 AI 产品",
];

const FRONT_COLUMN_TYPES: [ColumnType; FRONT_LEN] = [
    ColumnType::Text,                            // 实时清分UUID
    ColumnType::Text,                            // 商户号
    ColumnType::Text,                            // 商户名称
    ColumnType::Text,                            // 订单号
    ColumnType::Date,                            // 交易日期
    ColumnType::Decimal(DecimalScale::Original), // 交易金额
    ColumnType::Text,                            // 检索参考号
    ColumnType::Text,                            // 模版类型
    ColumnType::Text,                            // 状态
    ColumnType::Text,                            // 描述
    ColumnType::DateTime,                        // 提交时间
    ColumnType::DateTime,                        // 更新时间
    ColumnType::Text,                            // 终端号
    ColumnType::Text,                            // 分店id
    ColumnType::Text,                            // 分店名
    ColumnType::Text,                            // 所在地区
    ColumnType::Text,                            // 详细地址
    ColumnType::Text,                            // 地区编码
    ColumnType::Text,                            // tel
    ColumnType::Text,                            // 发票号码
    ColumnType::Text, // 发票金额（源值含“3000,00”等非标准写法，按文本保留）
    ColumnType::Text, // 购买方名称
    ColumnType::Text, // 图片1（第一次出现）
    ColumnType::Text, // S/N码
    ColumnType::Text, // 是否属于 AI 产品
];

/// 第26列起两组数据组结构不同，且已知存在同一数据组内部的表头变体（如“电脑”工作表），
/// 因此按名称在“第26列起”的范围内查找，不按固定列号；范围限定可避免与第23列同名的
/// `图片1`混淆。
pub(crate) struct TailField {
    name: &'static str,
    ty: ColumnType,
    /// 已确认的候选源字段名；第一项为该数据组的标准名称，用于测试生成标准表头。
    pub synonyms: &'static [&'static str],
    /// 是否为必需字段；`false`表示已知有工作表变体缺少该字段（如“电脑”工作表无
    /// `airConditionerKitInfo`），缺失时留空而非终止处理。
    required: bool,
}

/// 构造一条 `TailField`；未标注 `required: false` 时默认必需。
macro_rules! tail {
    ($name:expr, $ty:expr, [$($syn:expr),+ $(,)?]) => {
        tail!($name, $ty, [$($syn),+], true)
    };
    ($name:expr, $ty:expr, [$($syn:expr),+ $(,)?], $required:expr) => {
        TailField {
            name: $name,
            ty: $ty,
            synonyms: &[$($syn),+],
            required: $required,
        }
    };
}

pub(crate) const APPLIANCE_TAIL: [TailField; TAIL_LEN] = [
    tail!("图片1", ColumnType::Text, ["图片1"]),
    tail!("图片2", ColumnType::Text, ["图片2"]),
    tail!("图片3", ColumnType::Text, ["图片3"]),
    tail!("图片4", ColumnType::Text, ["图片4"]),
    // “电脑”工作表中该字段名为“图片5”，与“家电”工作表的“img5”同义。
    tail!("img5", ColumnType::Text, ["img5", "图片5"]),
    tail!("img6", ColumnType::Text, ["img6"]),
    tail!("img7", ColumnType::Text, ["img7"]),
    tail!("img8", ColumnType::Text, ["img8"]),
    tail!("img9", ColumnType::Text, ["img9"]),
    tail!("img10", ColumnType::Text, ["img10"]),
    tail!("img11", ColumnType::Text, ["img11"]),
    tail!("img12", ColumnType::Text, ["img12"]),
    tail!("img13", ColumnType::Text, ["img13"]),
    tail!("img14", ColumnType::Text, ["img14"]),
    tail!("img15", ColumnType::Text, ["img15"]),
    tail!("签收时间", ColumnType::DateTime, ["签收时间"]),
    tail!("remark", ColumnType::Text, ["remark"]),
    tail!("EEG", ColumnType::Text, ["EEG"]),
    tail!("物流单号", ColumnType::Text, ["物流单号"]),
    tail!("erpOrderNum", ColumnType::Text, ["erpOrderNum"]),
    tail!("ocrModify", ColumnType::Text, ["ocrModify"]),
    tail!("modifyStatus", ColumnType::Text, ["modifyStatus"]),
    tail!(
        "introduceInvoiceFlag",
        ColumnType::Text,
        ["introduceInvoiceFlag"]
    ),
    tail!("是否交旧", ColumnType::Text, ["是否交旧"]),
    tail!("是否自提", ColumnType::Text, ["是否自提"]),
    tail!("receiverName", ColumnType::Text, ["receiverName"]),
    tail!("productCode", ColumnType::Text, ["productCode"]),
    tail!("subsideAmt", ColumnType::Text, ["subsideAmt"]),
    tail!("productName", ColumnType::Text, ["productName"]),
    tail!("交旧品类", ColumnType::Text, ["交旧品类"]),
    tail!(
        "收货地址是否农村地区",
        ColumnType::Text,
        ["收货地址是否农村地区"]
    ),
    // “电脑”工作表没有此字段：缺失时留空，不终止处理。
    tail!(
        "airConditionerKitInfo",
        ColumnType::Text,
        ["airConditionerKitInfo"],
        false
    ),
    tail!("开票日期", ColumnType::Date, ["开票日期"]),
];

pub(crate) const DIGITAL_TAIL: [TailField; TAIL_LEN] = [
    tail!("IMEI1", ColumnType::Text, ["IMEI1"]),
    tail!("IMEI2", ColumnType::Text, ["IMEI2"]),
    tail!("图片1", ColumnType::Text, ["图片1"]),
    tail!("图片2", ColumnType::Text, ["图片2"]),
    tail!("图片3", ColumnType::Text, ["图片3"]),
    tail!("图片4", ColumnType::Text, ["图片4"]),
    tail!("图片5", ColumnType::Text, ["图片5"]),
    tail!("图片6", ColumnType::Text, ["图片6"]),
    tail!("img7", ColumnType::Text, ["img7"]),
    tail!("img8", ColumnType::Text, ["img8"]),
    tail!("img9", ColumnType::Text, ["img9"]),
    tail!("img10", ColumnType::Text, ["img10"]),
    tail!("img11", ColumnType::Text, ["img11"]),
    tail!("img12", ColumnType::Text, ["img12"]),
    tail!("img13", ColumnType::Text, ["img13"]),
    tail!("img14", ColumnType::Text, ["img14"]),
    tail!("img15", ColumnType::Text, ["img15"]),
    tail!("签收时间", ColumnType::DateTime, ["签收时间"]),
    tail!("remark", ColumnType::Text, ["remark"]),
    tail!("物流单号", ColumnType::Text, ["物流单号"]),
    tail!("erpOrderNum", ColumnType::Text, ["erpOrderNum"]),
    tail!("ocrModify", ColumnType::Text, ["ocrModify"]),
    tail!("modifyStatus", ColumnType::Text, ["modifyStatus"]),
    tail!(
        "introduceInvoiceFlag",
        ColumnType::Text,
        ["introduceInvoiceFlag"]
    ),
    tail!("是否交旧", ColumnType::Text, ["是否交旧"]),
    tail!("是否自提", ColumnType::Text, ["是否自提"]),
    tail!("receiverName", ColumnType::Text, ["receiverName"]),
    tail!("productCode", ColumnType::Text, ["productCode"]),
    tail!("subsideAmt", ColumnType::Text, ["subsideAmt"]),
    tail!("productName", ColumnType::Text, ["productName"]),
    tail!("oldExchangeType", ColumnType::Text, ["oldExchangeType"]),
    tail!(
        "收货地址是否农村地区",
        ColumnType::Text,
        ["收货地址是否农村地区"]
    ),
    tail!("开票日期", ColumnType::Date, ["开票日期"]),
];

// 固定列位置（1 基）。
const COL_UUID: u32 = 1;
const COL_MERCHANT_NO: u32 = 2;

struct UploadedConfig {
    category: Category,
    title: &'static str,
    output_stem: &'static str,
    merchant_no: &'static str,
    tail: &'static [TailField; TAIL_LEN],
}

const DIGITAL_CONFIG: UploadedConfig = UploadedConfig {
    category: Category::UploadedDigital,
    title: "数码已上传数据",
    output_stem: "已上传数码",
    merchant_no: "89813014812B06R",
    tail: &DIGITAL_TAIL,
};

const APPLIANCE_CONFIG: UploadedConfig = UploadedConfig {
    category: Category::UploadedAppliance,
    title: "家电、电脑已上传数据",
    output_stem: "已上传家电电脑",
    merchant_no: "89813015722APT1",
    tail: &APPLIANCE_TAIL,
};

/// 第5.7节：未命中回款明细时，“交易金额×15%”估算补贴金额的封顶值。
fn subsidy_cap(category: Category) -> Decimal {
    match category {
        Category::UploadedDigital => Decimal::new(50000, 2), // 500.00
        Category::UploadedAppliance => Decimal::new(150000, 2), // 1500.00
        other => unreachable!(
            "UploadedConfig 只用于 UploadedAppliance/UploadedDigital，实际为 {other:?}"
        ),
    }
}

pub struct UploadedJob(&'static UploadedConfig);

pub const UPLOADED_DIGITAL: UploadedJob = UploadedJob(&DIGITAL_CONFIG);
pub const UPLOADED_APPLIANCE: UploadedJob = UploadedJob(&APPLIANCE_CONFIG);

impl Job for UploadedJob {
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
        run_uploaded(self.0, input_dir)
    }
}

/// 文件名须匹配`MER_<商户号>_*.xlsx`。
fn matches_filename(name: &str, merchant_no: &str) -> bool {
    name.starts_with(&format!("MER_{merchant_no}_")) && name.ends_with(".xlsx")
}

/// 前25列须按固定列位置与名称完全一致；第26列起按名称在其范围内查找（第5.1节：
/// 数据组内部存在“电脑”等表头变体，不能仅按固定列号映射）。
fn resolve_columns(sheet: &SheetGrid, config: &UploadedConfig) -> Result<Vec<Option<u32>>, String> {
    let header = sheet.row_texts(2);
    if header.len() < FRONT_LEN
        || header[..FRONT_LEN]
            .iter()
            .map(String::as_str)
            .collect::<Vec<_>>()
            != FRONT_HEADERS
    {
        return Err(format!("第2行前{FRONT_LEN}列表头与规定字段及顺序不一致"));
    }

    let tail_domain = &header[FRONT_LEN..];
    let mut tail_columns = Vec::with_capacity(TAIL_LEN);
    for field in config.tail {
        let resolved =
            resolve_synonym_column(tail_domain, field.synonyms)?.map(|col| FRONT_LEN as u32 + col);
        if resolved.is_none() && field.required {
            return Err(format!("缺少必需字段：{}", field.name));
        }
        tail_columns.push(resolved);
    }
    Ok(tail_columns)
}

fn read_typed_cell(
    cell: &RawCell,
    ty: ColumnType,
    field: &str,
    file: &str,
    sheet_name: &str,
    row: u32,
) -> Result<Value, ProcessError> {
    Ok(match ty {
        ColumnType::Date => cell_date_or_text(cell),
        ColumnType::DateTime => cell_datetime_or_text(cell),
        ColumnType::Decimal(_) => cell_amount(cell)
            .map(amount_value)
            .map_err(|detail| data_error(file, sheet_name, row, field, cell.to_string(), detail))?,
        _ => text_value(cell_text(cell).map_err(|detail| {
            data_error(file, sheet_name, row, field, cell.to_string(), detail)
        })?),
    })
}

fn read_row(
    sheet: &SheetGrid,
    row: u32,
    config: &UploadedConfig,
    tail_columns: &[Option<u32>],
    file: &str,
    sheet_name: &str,
) -> Result<Row, ProcessError> {
    let merchant_cell = sheet.cell(row, COL_MERCHANT_NO);
    let merchant_no = cell_text(&merchant_cell).map_err(|detail| {
        data_error(
            file,
            sheet_name,
            row,
            "商户号",
            merchant_cell.to_string(),
            detail,
        )
    })?;
    if merchant_no != config.merchant_no {
        return Err(data_error(
            file,
            sheet_name,
            row,
            "商户号",
            merchant_no,
            format!("与文件名商户号“{}”不一致", config.merchant_no),
        ));
    }

    let mut values = Vec::with_capacity(FIELD_COUNT);
    for (index, &ty) in FRONT_COLUMN_TYPES.iter().enumerate() {
        let col = (index + 1) as u32;
        let cell = sheet.cell(row, col);
        values.push(read_typed_cell(
            &cell,
            ty,
            FRONT_HEADERS[index],
            file,
            sheet_name,
            row,
        )?);
    }
    for (index, field) in config.tail.iter().enumerate() {
        let value = match tail_columns[index] {
            None => Value::Empty,
            Some(col) => {
                let cell = sheet.cell(row, col);
                read_typed_cell(&cell, field.ty, field.name, file, sheet_name, row)?
            }
        };
        values.push(value);
    }

    Ok(Row { values, fill: None })
}

fn output_columns(config: &UploadedConfig) -> Vec<Column> {
    FRONT_HEADERS
        .iter()
        .zip(FRONT_COLUMN_TYPES)
        .map(|(&name, ty)| Column { name, ty })
        .chain(config.tail.iter().map(|field| Column {
            name: field.name,
            ty: field.ty,
        }))
        .chain(std::iter::once(Column {
            name: "补贴金额",
            ty: ColumnType::Decimal(DecimalScale::Two),
        }))
        .collect()
}

// 已上传数据前25列固定结构中用于匹配/覆盖的字段位置（0 基，对齐 Row.values）。
const VALUE_COL_ORDER_NO: usize = 3; // 订单号
const VALUE_COL_TRANSACTION_AMOUNT: usize = 5; // 交易金额
const VALUE_COL_REFERENCE: usize = 6; // 检索参考号
const VALUE_COL_STATUS: usize = 8; // 状态
const VALUE_COL_INVOICE_NO: usize = 19; // 发票号码

// 回款明细（refund.rs）24 列统一输出中用于匹配的字段位置（0 基）。
const REFUND_COL_REFERENCE: usize = 2; // 交易参考号
const REFUND_COL_MERCHANT_ORDER: usize = 3; // 商户订单号
const REFUND_COL_SUBSIDY: usize = 10; // 补贴金额
const REFUND_COL_INVOICE_NO: usize = 19; // 发票号

/// 按`商户订单号`、`交易参考号`、`发票号`分别汇总回款明细"正常区域"（排除第7.5/8.5节
/// 沉底的粉色重复记录）全部`补贴金额`（未去重，格式化为两位小数文本以统一比较）。
fn build_refund_indices(
    refund_table: &Table,
) -> (MultiValueIndex, MultiValueIndex, MultiValueIndex) {
    let mut by_order: MultiValueIndex = HashMap::new();
    let mut by_reference: MultiValueIndex = HashMap::new();
    let mut by_invoice_no: MultiValueIndex = HashMap::new();

    for row in refund_table
        .rows
        .iter()
        .filter(|r| r.fill != Some(Fill::Pink))
    {
        let Value::Decimal(subsidy) = &row.values[REFUND_COL_SUBSIDY] else {
            continue;
        };
        let subsidy_text = subsidy.round_dp(2).to_string();

        if let Value::Text(order_no) = &row.values[REFUND_COL_MERCHANT_ORDER] {
            by_order
                .entry(order_no.clone())
                .or_default()
                .push(subsidy_text.clone());
        }
        if let Value::Text(reference) = &row.values[REFUND_COL_REFERENCE] {
            by_reference
                .entry(reference.clone())
                .or_default()
                .push(subsidy_text.clone());
        }
        if let Value::Text(invoice_no) = &row.values[REFUND_COL_INVOICE_NO] {
            by_invoice_no
                .entry(invoice_no.clone())
                .or_default()
                .push(subsidy_text);
        }
    }
    (by_order, by_reference, by_invoice_no)
}

/// 按`订单号`（主键）→`检索参考号`（次键）→`发票号码`（三键）依次在回款明细索引中查找
/// 唯一命中的`补贴金额`；命中即停，同级歧义立即停止、不再尝试下一级。
fn matched_refund_subsidy(
    by_order: &MultiValueIndex,
    by_reference: &MultiValueIndex,
    by_invoice_no: &MultiValueIndex,
    order_no: &str,
    reference: &str,
    invoice_no: &str,
) -> Option<Decimal> {
    for (index, key) in [
        (by_order, order_no),
        (by_reference, reference),
        (by_invoice_no, invoice_no),
    ] {
        match resolve_via(index, Some(key)) {
            PriorityOutcome::Unique(value) => return Decimal::from_str(&value).ok(),
            PriorityOutcome::Ambiguous => return None,
            PriorityOutcome::NoHit => {}
        }
    }
    None
}

fn value_text(value: &Value) -> &str {
    match value {
        Value::Text(s) => s.as_str(),
        _ => "",
    }
}

/// 第5.7节：未命中回款明细时，按“交易金额×15%”估算补贴金额并按`cap`封顶；
/// 交易金额缺失或非数值时无法估算，返回`None`（补贴金额留空，不视为命中回款明细）。
fn estimate_subsidy(transaction_amount: &Value, cap: Decimal) -> Option<Decimal> {
    let Value::Decimal(amount) = transaction_amount else {
        return None;
    };
    let ratio = Decimal::new(15, 2); // 0.15
    Some((*amount * ratio).round_dp(2).min(cap))
}

/// 第5.7节状态归一化：仅替换列出的原值，其余原样保留；在回款匹配的“已回款”覆盖之前
/// 执行，命中回款明细时“已回款”仍会覆盖本函数的归并结果。
fn normalize_status(value: String) -> String {
    match value.as_str() {
        "审核通过" => "审核通过未回款".to_string(),
        "同步(已上送)" | "暂存" | "待同步" => "待审核".to_string(),
        "核销失败" => "审核失败".to_string(),
        _ => value,
    }
}

fn run_uploaded(config: &UploadedConfig, input_dir: &Path) -> Result<Table, ProcessError> {
    let mut files: Vec<PathBuf> = list_xlsx_files(input_dir)?
        .into_iter()
        .filter(|path| {
            path.file_name()
                .and_then(|n| n.to_str())
                .is_some_and(|n| matches_filename(n, config.merchant_no))
        })
        .collect();
    files.sort();
    if files.is_empty() {
        return Err(ProcessError::NoInput {
            pattern: format!("MER_{}_*.xlsx", config.merchant_no),
        });
    }

    // 补贴金额匹配的回款明细依赖：缺失或结构异常时同样必须停止，直接复用 RefundJob::run()
    // 的全部校验；只取其“正常区域”（非粉色沉底/重复记录）作为匹配对象。
    let refund_table = match config.category {
        Category::UploadedAppliance => REFUND_APPLIANCE.run(input_dir),
        Category::UploadedDigital => REFUND_DIGITAL.run(input_dir),
        other => unreachable!(
            "UploadedConfig 只用于 UploadedAppliance/UploadedDigital，实际为 {other:?}"
        ),
    }?;
    let (by_order, by_reference, by_invoice_no) = build_refund_indices(&refund_table);
    let subsidy_cap = subsidy_cap(config.category);

    let mut rows = Vec::new();
    let mut fingerprints: HashMap<String, String> = HashMap::new();
    let mut seen_uuid: HashSet<String> = HashSet::new();

    for path in &files {
        let file_name = path.file_name().unwrap().to_string_lossy().into_owned();
        let sheets = open_sheets(path)?;

        for sheet in &sheets {
            let sheet_name = sheet.name().to_string();

            let tail_columns =
                resolve_columns(sheet, config).map_err(|detail| ProcessError::Structure {
                    file: file_name.clone(),
                    sheet: sheet_name.clone(),
                    detail,
                })?;

            let last_row = sheet.last_value_row().unwrap_or(2);

            let fingerprint = sheet.fingerprint(1, last_row);
            check_duplicate_fingerprint(&mut fingerprints, fingerprint, &file_name)?;

            for row in 3..=last_row {
                let mut record_row =
                    read_row(sheet, row, config, &tail_columns, &file_name, &sheet_name)?;
                if let Value::Text(uuid) = &record_row.values[(COL_UUID - 1) as usize]
                    && !seen_uuid.insert(uuid.clone())
                {
                    return Err(ProcessError::Duplicate {
                        detail: format!("实时清分UUID重复：{uuid}"),
                    });
                }

                if let Value::Text(status) = &record_row.values[VALUE_COL_STATUS] {
                    record_row.values[VALUE_COL_STATUS] =
                        Value::Text(normalize_status(status.clone()));
                }

                let subsidy = matched_refund_subsidy(
                    &by_order,
                    &by_reference,
                    &by_invoice_no,
                    value_text(&record_row.values[VALUE_COL_ORDER_NO]),
                    value_text(&record_row.values[VALUE_COL_REFERENCE]),
                    value_text(&record_row.values[VALUE_COL_INVOICE_NO]),
                );
                if subsidy.is_some() {
                    record_row.values[VALUE_COL_STATUS] = Value::Text("已回款".to_string());
                }
                // 未命中回款明细时，按“交易金额×15%”估算并封顶（第5.7节）；估算不改变
                // 状态，只补充补贴金额，避免与真正命中回款明细的“已回款”混淆。
                let subsidy = subsidy.or_else(|| {
                    estimate_subsidy(
                        &record_row.values[VALUE_COL_TRANSACTION_AMOUNT],
                        subsidy_cap,
                    )
                });
                record_row
                    .values
                    .push(subsidy.map_or(Value::Empty, Value::Decimal));

                rows.push(record_row);
            }
        }
    }

    Ok(Table {
        columns: output_columns(config),
        rows,
    })
}

#[cfg(test)]
mod tests {
    use rust_xlsxwriter::Workbook;

    use super::*;
    use crate::test_support::unique_temp_path;

    const REFUND_APPLIANCE_HEADER: [&str; 19] = [
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

    const REFUND_DIGITAL_HEADER: [&str; 13] = [
        "拨付批次",
        "参考号",
        "商户订单号",
        "销方名称",
        "经销商编号",
        "应收销售金额",
        "补贴销售金额",
        "SN码",
        "所在地区",
        "商品编码",
        "品类",
        "商品明细",
        "发票号码",
    ];

    /// 构造一份最小合法的回款明细家电电脑样本（单工作表，第1行表头、第2行起明细），
    /// 供补贴金额匹配测试；`rows`为`(交易参考号, 商户订单号, 发票号, 补贴金额)`四元组。
    fn write_refund_appliance_fixture(dir: &Path, rows: &[(&str, &str, &str, &str)]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        for (col, name) in REFUND_APPLIANCE_HEADER.iter().enumerate() {
            sheet.write_string(0, col as u16, *name).unwrap();
        }
        for (index, (reference, merchant_order, invoice_no, subsidy)) in rows.iter().enumerate() {
            let row = (1 + index) as u32;
            let values = [
                "批次1",
                reference,
                merchant_order,
                "",
                "企业甲",
                "89813015722APT1",
                "0.00",
                "100.00",
                "90.00",
                subsidy,
                "0.15",
                "SN001",
                "地区甲",
                "CODE001",
                "一级",
                "品类甲",
                "商品甲",
                invoice_no,
                "ID001",
            ];
            for (col, value) in values.iter().enumerate() {
                sheet.write_string(row, col as u16, *value).unwrap();
            }
        }
        workbook
            .save(dir.join("2026年以旧换新补贴明细.xlsx"))
            .unwrap();
    }

    /// 构造一份最小合法的回款明细数码样本；`rows`为`(交易参考号, 商户订单号, 发票号, 补贴金额)`四元组。
    fn write_refund_digital_fixture(dir: &Path, rows: &[(&str, &str, &str, &str)]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        for (col, name) in REFUND_DIGITAL_HEADER.iter().enumerate() {
            sheet.write_string(0, col as u16, *name).unwrap();
        }
        for (index, (reference, merchant_order, invoice_no, subsidy)) in rows.iter().enumerate() {
            let row = (1 + index) as u32;
            let values = [
                "批次1",
                reference,
                merchant_order,
                "企业乙",
                "89813014812B06R",
                "100.00",
                subsidy,
                "SN002",
                "地区乙",
                "CODE002",
                "品类乙",
                "商品乙",
                invoice_no,
            ];
            for (col, value) in values.iter().enumerate() {
                sheet.write_string(row, col as u16, *value).unwrap();
            }
        }
        workbook.save(dir.join("2026年数码补贴明细.xlsx")).unwrap();
    }

    fn write_workbook(path: &Path, title: &str, headers: &[&str], rows: &[Vec<&str>]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, title).unwrap();
        for (col, name) in headers.iter().enumerate() {
            sheet.write_string(1, col as u16, *name).unwrap();
        }
        for (row_index, row) in rows.iter().enumerate() {
            for (col, value) in row.iter().enumerate() {
                sheet
                    .write_string((2 + row_index) as u32, col as u16, *value)
                    .unwrap();
            }
        }
        workbook.save(path).unwrap();
    }

    /// 按数据组标准表头（前25列固定 + 该组的尾部字段名）生成表头。
    fn standard_headers(tail: &[TailField]) -> Vec<&'static str> {
        FRONT_HEADERS
            .iter()
            .copied()
            .chain(tail.iter().map(|f| f.synonyms[0]))
            .collect()
    }

    /// 按`APPLIANCE_TAIL`顺序生成一行示例数据；两个`图片1`用不同值验证按位置区分。
    fn appliance_row<'a>(uuid: &'a str, merchant_no: &'a str, order_no: &'a str) -> Vec<&'a str> {
        vec![
            uuid,
            merchant_no,
            "商户甲",
            order_no,
            "20260107",
            "100.00",
            "REF001",
            "模版A",
            "已上传",
            "描述文本",
            "2026-09-14 10:18:09",
            "2026-09-14 10:20:00",
            "T001",
            "STORE001",
            "分店甲",
            "地区甲",
            "详细地址甲",
            "110000",
            "13800000000",
            "INV001",
            "3000,00",
            "购买方甲",
            "PIC_FIRST",
            "SN0001",
            "是",
            "PIC_SECOND",
            "PIC2",
            "PIC3",
            "PIC4",
            "IMG5",
            "IMG6",
            "IMG7",
            "IMG8",
            "IMG9",
            "IMG10",
            "IMG11",
            "IMG12",
            "IMG13",
            "IMG14",
            "IMG15",
            "2026-09-14 12:00:00",
            "备注甲",
            "EEG甲",
            "LOG001",
            "ERP001",
            "ocr甲",
            "状态甲",
            "标记甲",
            "是",
            "否",
            "收件人甲",
            "PC001",
            "",
            "商品甲",
            "品类甲",
            "否",
            "空调信息甲",
            "20260109",
        ]
    }

    /// 按`DIGITAL_TAIL`顺序生成一行示例数据。
    fn digital_row<'a>(uuid: &'a str, merchant_no: &'a str, order_no: &'a str) -> Vec<&'a str> {
        vec![
            uuid,
            merchant_no,
            "商户乙",
            order_no,
            "20260108",
            "200.00",
            "REF002",
            "模版B",
            "已上传",
            "描述文本2",
            "2026-09-15 11:00:00",
            "2026-09-15 11:05:00",
            "T002",
            "STORE002",
            "分店乙",
            "地区乙",
            "详细地址乙",
            "220000",
            "13900000000",
            "INV002",
            "6000,00",
            "购买方乙",
            "PIC_FIRST2",
            "SN0002",
            "否",
            "IMEI0001",
            "IMEI0002",
            "PIC_SECOND2",
            "PIC2",
            "PIC3",
            "PIC4",
            "PIC5",
            "PIC6",
            "IMG7",
            "IMG8",
            "IMG9",
            "IMG10",
            "IMG11",
            "IMG12",
            "IMG13",
            "IMG14",
            "IMG15",
            "2026-09-15 13:00:00",
            "备注乙",
            "LOG002",
            "ERP002",
            "ocr乙",
            "状态乙",
            "标记乙",
            "否",
            "是",
            "收件人乙",
            "PC002",
            "",
            "商品乙",
            "旧换类型乙",
            "是",
            "20260110",
        ]
    }

    #[test]
    fn matches_filename_checks_merchant_prefix_and_extension() {
        assert!(matches_filename(
            "MER_89813014812B06R_20260914101809_yjhx.xlsx",
            "89813014812B06R"
        ));
        assert!(matches_filename(
            "MER_89813014812B06R_anything.xlsx",
            "89813014812B06R"
        ));
        assert!(!matches_filename(
            "MER_89813015722APT1_20260914101809_yjhx.xlsx",
            "89813014812B06R"
        ));
        assert!(!matches_filename(
            "MER_89813014812B06R_20260914101809_yjhx.xls",
            "89813014812B06R"
        ));
    }

    #[test]
    fn digital_job_reads_59_columns_and_parses_typed_fields() {
        let dir = unique_temp_path("uploaded-digital-happy-path");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_digital_fixture(&dir, &[]); // 回款明细无命中，补贴金额留空
        write_workbook(
            &dir.join("MER_89813014812B06R_20260914101809_yjhx.xlsx"),
            "以旧换新数据[手机/其他3C]",
            &standard_headers(&DIGITAL_TAIL),
            &[digital_row("UUID-1", "89813014812B06R", "ORDER-1")],
        );

        let table = UPLOADED_DIGITAL.run(&dir).unwrap();

        assert_eq!(table.columns.len(), 59);
        assert_eq!(table.rows.len(), 1);
        let values = &table.rows[0].values;
        assert_eq!(values[0], Value::Text("UUID-1".to_string()));
        assert!(matches!(values[4], Value::Date(_))); // 交易日期
        assert!(matches!(values[10], Value::DateTime(_))); // 提交时间
        assert!(matches!(values[42], Value::DateTime(_))); // 签收时间
        assert!(matches!(values[57], Value::Date(_))); // 开票日期
        assert_eq!(values[20], Value::Text("6000,00".to_string())); // 发票金额：原样保留，不解析为数值
        assert_eq!(values[25], Value::Text("IMEI0001".to_string()));
        assert_eq!(values[26], Value::Text("IMEI0002".to_string()));
        // 两个“图片1”按位置区分，取值不同。
        assert_eq!(values[22], Value::Text("PIC_FIRST2".to_string()));
        assert_eq!(values[27], Value::Text("PIC_SECOND2".to_string()));
        assert_eq!(values[8], Value::Text("已上传".to_string())); // 未命中回款明细，状态保持原值
        // 补贴金额：未命中时按“交易金额×15%”估算；digital_row 交易金额为 200.00 → 30.00。
        assert_eq!(values[58], Value::Decimal("30.00".parse().unwrap()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn appliance_job_has_59_columns_and_ignores_digital_files() {
        let dir = unique_temp_path("uploaded-appliance-happy-path");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[]); // 回款明细无命中，补贴金额留空
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-2", "89813015722APT1", "ORDER-2")],
        );
        // 数码商户的文件必须被家电、电脑数据组忽略。
        write_workbook(
            &dir.join("MER_89813014812B06R_20260914101809_yjhx.xlsx"),
            "数码",
            &standard_headers(&DIGITAL_TAIL),
            &[digital_row("UUID-3", "89813014812B06R", "ORDER-3")],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();

        assert_eq!(table.columns.len(), 59);
        assert_eq!(table.rows.len(), 1);
        let values = &table.rows[0].values;
        assert_eq!(values[0], Value::Text("UUID-2".to_string()));
        assert_eq!(values[2], Value::Text("商户甲".to_string())); // 商户名称
        assert!(matches!(values[40], Value::DateTime(_))); // 签收时间（家电结构无 IMEI，位置在40）
        assert_eq!(values[20], Value::Text("3000,00".to_string())); // 发票金额
        assert_eq!(values[22], Value::Text("PIC_FIRST".to_string()));
        assert_eq!(values[25], Value::Text("PIC_SECOND".to_string()));
        // 补贴金额：未命中时按“交易金额×15%”估算；appliance_row 交易金额为 100.00 → 15.00。
        assert_eq!(values[58], Value::Decimal("15.00".parse().unwrap()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    /// 已知的“电脑”工作表变体：第26列多出`电脑类型`（不输出）、第5张图片改名为`图片5`
    /// （而非`img5`），且没有`airConditionerKitInfo`。必须仍能按名称正确映射到统一输出
    /// 结构，缺失字段留空而不是终止处理。
    #[test]
    fn maps_computer_variant_sheet_by_name_and_leaves_missing_field_blank() {
        let dir = unique_temp_path("uploaded-computer-variant");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[]);

        let mut headers = standard_headers(&APPLIANCE_TAIL);
        headers.insert(FRONT_LEN, "电脑类型"); // 不属于统一表头的额外字段
        let img5_pos = headers.iter().position(|h| *h == "img5").unwrap();
        headers[img5_pos] = "图片5"; // “电脑”表使用的同义名称
        let ac_pos = headers
            .iter()
            .position(|h| *h == "airConditionerKitInfo")
            .unwrap();
        headers.remove(ac_pos); // “电脑”表没有此字段

        // 行数据与表头做相同的插入/删除，值随其原字段一起移动，位置始终对应正确的表头。
        let mut row = appliance_row("UUID-PC", "89813015722APT1", "ORDER-PC");
        row.insert(FRONT_LEN, "电脑类型值");
        row.remove(ac_pos);

        write_workbook(
            &dir.join("MER_89813015722APT1_20260914102100_yjhx.xlsx"),
            "电脑",
            &headers,
            &[row],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(table.rows.len(), 1);
        let values = &table.rows[0].values;
        assert_eq!(values[0], Value::Text("UUID-PC".to_string()));
        assert_eq!(values[29], Value::Text("IMG5".to_string())); // img5 通过同义名称“图片5”映射
        assert_eq!(values[56], Value::Empty); // airConditionerKitInfo 缺失，留空
        assert!(matches!(values[57], Value::Date(_))); // 开票日期仍正确映射

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_merchant_mismatch_between_filename_and_row() {
        let dir = unique_temp_path("uploaded-merchant-mismatch");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[]);
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-5", "89813014812B06R", "ORDER-5")],
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Data { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_front_header_that_does_not_match_fixed_structure() {
        let dir = unique_temp_path("uploaded-header-mismatch");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[]);
        let mut headers = standard_headers(&APPLIANCE_TAIL);
        headers[13] = "门店编号"; // 与规定的“分店id”不一致
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &headers,
            &[appliance_row("UUID-6", "89813015722APT1", "ORDER-6")],
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Structure { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_duplicate_uuid_across_files() {
        let dir = unique_temp_path("uploaded-duplicate-uuid");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[]);
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-SAME", "89813015722APT1", "ORDER-8")],
        );
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101900_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-SAME", "89813015722APT1", "ORDER-9")],
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Duplicate { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_whole_sheet_duplicate_export_across_files() {
        let dir = unique_temp_path("uploaded-duplicate-sheet");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[]);
        let rows = [appliance_row("UUID-DUP", "89813015722APT1", "ORDER-10")];
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &rows,
        );
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101900_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &rows,
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Duplicate { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn unparseable_date_falls_back_to_original_text() {
        let dir = unique_temp_path("uploaded-bad-date");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[]);
        let mut row = appliance_row("UUID-11", "89813015722APT1", "ORDER-11");
        row[4] = "不是日期";

        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[row],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[4], Value::Text("不是日期".to_string()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_nothing_matches() {
        let dir = unique_temp_path("uploaded-no-input");
        std::fs::create_dir_all(&dir).unwrap();

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_refund_source_missing() {
        let dir = unique_temp_path("uploaded-missing-refund");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row(
                "UUID-NOREFUND",
                "89813015722APT1",
                "ORDER-NOREFUND",
            )],
        );

        // 回款明细家电电脑样本文件缺失：必须停止，不得让补贴金额列全部留空后继续输出。
        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn normalizes_listed_status_values_and_keeps_others() {
        assert_eq!(normalize_status("审核通过".to_string()), "审核通过未回款");
        for original in ["同步(已上送)", "暂存", "待同步"] {
            assert_eq!(normalize_status(original.to_string()), "待审核");
        }
        assert_eq!(normalize_status("核销失败".to_string()), "审核失败");
        assert_eq!(normalize_status("待审核".to_string()), "待审核"); // 未列出，原样保留
        assert_eq!(normalize_status("审核失败".to_string()), "审核失败"); // 未列出，原样保留
    }

    #[test]
    fn unmatched_row_gets_normalized_status() {
        let dir = unique_temp_path("uploaded-status-normalize-unmatched");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[]); // 无命中
        let mut row = appliance_row("UUID-G", "89813015722APT1", "ORDER-G");
        row[VALUE_COL_STATUS] = "核销失败";
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[row],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(
            table.rows[0].values[VALUE_COL_STATUS],
            Value::Text("审核失败".to_string())
        );
        // 未命中回款明细，按“交易金额×15%”估算补贴金额；appliance_row 交易金额为 100.00 → 15.00。
        assert_eq!(
            table.rows[0].values[58],
            Value::Decimal("15.00".parse().unwrap())
        );

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn refund_match_overrides_normalized_status_with_already_repaid() {
        let dir = unique_temp_path("uploaded-status-normalize-then-match");
        std::fs::create_dir_all(&dir).unwrap();
        // 命中回款明细时“已回款”须覆盖归并结果，即使原状态在归并表中。
        write_refund_appliance_fixture(&dir, &[("NOMATCH-REF", "ORDER-H", "NOMATCH-INV", "9.99")]);
        let mut row = appliance_row("UUID-H", "89813015722APT1", "ORDER-H");
        row[VALUE_COL_STATUS] = "审核通过";
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[row],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(
            table.rows[0].values[VALUE_COL_STATUS],
            Value::Text("已回款".to_string())
        );
        assert_eq!(
            table.rows[0].values[58],
            Value::Decimal("9.99".parse().unwrap())
        );

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn estimate_subsidy_applies_ratio_and_cap() {
        let amount = Value::Decimal("1000.00".parse().unwrap());
        assert_eq!(
            estimate_subsidy(&amount, Decimal::new(150000, 2)),
            Some("150.00".parse().unwrap())
        );
        let big_amount = Value::Decimal("20000.00".parse().unwrap());
        assert_eq!(
            estimate_subsidy(&big_amount, Decimal::new(150000, 2)),
            Some("1500.00".parse().unwrap())
        );
        assert_eq!(
            estimate_subsidy(&Value::Empty, Decimal::new(150000, 2)),
            None
        );
    }

    #[test]
    fn caps_estimated_subsidy_at_appliance_limit() {
        let dir = unique_temp_path("uploaded-refund-estimate-cap-appliance");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[]); // 无命中，走估算
        let mut row = appliance_row("UUID-I", "89813015722APT1", "ORDER-I");
        row[VALUE_COL_TRANSACTION_AMOUNT] = "20000.00"; // ×15%=3000.00，超过家电电脑封顶1500
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[row],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(
            table.rows[0].values[58],
            Value::Decimal("1500.00".parse().unwrap())
        );

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn caps_estimated_subsidy_at_digital_limit() {
        let dir = unique_temp_path("uploaded-refund-estimate-cap-digital");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_digital_fixture(&dir, &[]); // 无命中，走估算
        let mut row = digital_row("UUID-J", "89813014812B06R", "ORDER-J");
        row[VALUE_COL_TRANSACTION_AMOUNT] = "10000.00"; // ×15%=1500.00，超过数码封顶500
        write_workbook(
            &dir.join("MER_89813014812B06R_20260914101809_yjhx.xlsx"),
            "数码",
            &standard_headers(&DIGITAL_TAIL),
            &[row],
        );

        let table = UPLOADED_DIGITAL.run(&dir).unwrap();
        assert_eq!(
            table.rows[0].values[58],
            Value::Decimal("500.00".parse().unwrap())
        );

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn leaves_subsidy_empty_when_transaction_amount_missing_and_unmatched() {
        let dir = unique_temp_path("uploaded-refund-estimate-missing-amount");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[]); // 无命中
        let mut row = appliance_row("UUID-K", "89813015722APT1", "ORDER-K");
        row[VALUE_COL_TRANSACTION_AMOUNT] = ""; // 交易金额缺失，无法估算
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[row],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[58], Value::Empty);

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn order_no_match_replaces_status_and_fills_subsidy_amount() {
        let dir = unique_temp_path("uploaded-refund-order-hit");
        std::fs::create_dir_all(&dir).unwrap();
        write_refund_appliance_fixture(&dir, &[("NOMATCH-REF", "ORDER-A", "NOMATCH-INV", "88.88")]);
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-A", "89813015722APT1", "ORDER-A")],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        let values = &table.rows[0].values;
        assert_eq!(values[VALUE_COL_STATUS], Value::Text("已回款".to_string()));
        assert_eq!(values[58], Value::Decimal("88.88".parse().unwrap()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn falls_back_to_reference_when_order_no_has_no_hit() {
        let dir = unique_temp_path("uploaded-refund-reference-fallback");
        std::fs::create_dir_all(&dir).unwrap();
        // 商户订单号在回款明细中查无，须降级到检索参考号（次键）；REF001 是 appliance_row
        // 固定写入的检索参考号。
        write_refund_appliance_fixture(
            &dir,
            &[("REF001", "NOMATCH-ORDER", "NOMATCH-INV", "12.34")],
        );
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-B", "89813015722APT1", "ORDER-B")],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        let values = &table.rows[0].values;
        assert_eq!(values[VALUE_COL_STATUS], Value::Text("已回款".to_string()));
        assert_eq!(values[58], Value::Decimal("12.34".parse().unwrap()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn falls_back_to_invoice_no_when_order_no_and_reference_have_no_hit() {
        let dir = unique_temp_path("uploaded-refund-invoice-fallback");
        std::fs::create_dir_all(&dir).unwrap();
        // 订单号和检索参考号均查无，须降级到发票号码（三键）；INV001 是固定写入的发票号码。
        write_refund_appliance_fixture(
            &dir,
            &[("NOMATCH-REF", "NOMATCH-ORDER", "INV001", "56.70")],
        );
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-C", "89813015722APT1", "ORDER-C")],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        let values = &table.rows[0].values;
        assert_eq!(values[VALUE_COL_STATUS], Value::Text("已回款".to_string()));
        assert_eq!(values[58], Value::Decimal("56.70".parse().unwrap()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn order_no_match_takes_priority_over_reference_and_invoice_no() {
        let dir = unique_temp_path("uploaded-refund-priority");
        std::fs::create_dir_all(&dir).unwrap();
        // 三个键各自命中不同的补贴金额：主键（订单号）命中的值必须胜出。
        write_refund_appliance_fixture(
            &dir,
            &[
                ("NOMATCH-REF", "ORDER-D", "NOMATCH-INV", "1.00"), // 订单号命中
                ("REF001", "NOMATCH-ORDER-1", "NOMATCH-INV-1", "2.00"), // 检索参考号命中
                ("NOMATCH-REF-2", "NOMATCH-ORDER-2", "INV001", "3.00"), // 发票号码命中
            ],
        );
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-D", "89813015722APT1", "ORDER-D")],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(
            table.rows[0].values[58],
            Value::Decimal("1.00".parse().unwrap())
        );

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn ambiguous_order_no_match_stops_without_falling_through() {
        let dir = unique_temp_path("uploaded-refund-ambiguous");
        std::fs::create_dir_all(&dir).unwrap();
        // 两条回款明细各自的交易订单号不同（因此不属于同一回款内部重复分组、均保留在
        // 正常区域），但商户订单号相同、补贴金额不同：对本任务的匹配索引而言构成歧义，
        // 须立即判定未命中，不得降级到检索参考号（即使 REF001 本可命中第三条记录）。
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        for (col, name) in REFUND_APPLIANCE_HEADER.iter().enumerate() {
            sheet.write_string(0, col as u16, *name).unwrap();
        }
        let rows: [[&str; 19]; 3] = [
            [
                "批次1",
                "NOMATCH-REF-1",
                "ORDER-E",
                "TXN-1",
                "企业甲",
                "89813015722APT1",
                "0.00",
                "100.00",
                "90.00",
                "1.00",
                "0.15",
                "SN001",
                "地区甲",
                "CODE001",
                "一级",
                "品类甲",
                "商品甲",
                "NOMATCH-INV-1",
                "ID001",
            ],
            [
                "批次1",
                "NOMATCH-REF-2",
                "ORDER-E",
                "TXN-2",
                "企业甲",
                "89813015722APT1",
                "0.00",
                "100.00",
                "90.00",
                "2.00",
                "0.15",
                "SN001",
                "地区甲",
                "CODE001",
                "一级",
                "品类甲",
                "商品甲",
                "NOMATCH-INV-2",
                "ID001",
            ],
            [
                "批次1",
                "REF001",
                "NOMATCH-ORDER",
                "TXN-3",
                "企业甲",
                "89813015722APT1",
                "0.00",
                "100.00",
                "90.00",
                "3.00",
                "0.15",
                "SN001",
                "地区甲",
                "CODE001",
                "一级",
                "品类甲",
                "商品甲",
                "NOMATCH-INV-3",
                "ID001",
            ],
        ];
        for (index, row) in rows.iter().enumerate() {
            let excel_row = (1 + index) as u32;
            for (col, value) in row.iter().enumerate() {
                sheet.write_string(excel_row, col as u16, *value).unwrap();
            }
        }
        workbook
            .save(dir.join("2026年以旧换新补贴明细.xlsx"))
            .unwrap();
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-E", "89813015722APT1", "ORDER-E")],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        let values = &table.rows[0].values;
        assert_eq!(values[VALUE_COL_STATUS], Value::Text("已上传".to_string())); // 状态保持原值
        // 歧义按未命中处理，按“交易金额×15%”估算补贴金额；appliance_row 交易金额为 100.00 → 15.00。
        assert_eq!(values[58], Value::Decimal("15.00".parse().unwrap()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn excludes_pink_sunk_refund_records_from_matching_pool() {
        let dir = unique_temp_path("uploaded-refund-pink-excluded");
        std::fs::create_dir_all(&dir).unwrap();
        // 两条回款明细共享同一商户订单号（交易订单号留空，按7.5.1节分组依据降级到商户
        // 订单号）；补贴金额合计为0，按7.5.2节整组沉底并标记粉色，不得进入匹配索引。
        write_refund_appliance_fixture(
            &dir,
            &[
                ("NOMATCH-REF-1", "ORDER-F", "NOMATCH-INV-1", "10.00"),
                ("NOMATCH-REF-2", "ORDER-F", "NOMATCH-INV-2", "-10.00"),
            ],
        );
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-F", "89813015722APT1", "ORDER-F")],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        let values = &table.rows[0].values;
        assert_eq!(values[VALUE_COL_STATUS], Value::Text("已上传".to_string())); // 未命中，状态保持原值
        // 沉底记录不参与匹配，按“交易金额×15%”估算补贴金额；appliance_row 交易金额为 100.00 → 15.00。
        assert_eq!(values[58], Value::Decimal("15.00".parse().unwrap()));

        std::fs::remove_dir_all(&dir).unwrap();
    }
}
