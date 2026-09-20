use std::collections::{HashMap, HashSet};
use std::path::Path;

use rust_decimal::Decimal;

use crate::io::paths::list_xlsx_files;
use crate::io::xlsx_reader::{RawCell, SheetGrid, open_sheets};
use crate::model::{Column, ColumnType, DecimalScale, Fill, ProcessError, Row, Table, Value};

use super::{
    Category, Job, amount_value, cell_amount, cell_date_or_text, cell_datetime_or_text, cell_text,
    check_duplicate_fingerprint, data_error, text_value,
};

pub(crate) const HEADERS: [&str; 26] = [
    "清算时间",
    "交易时间",
    "终端号",
    "交易类型",
    "卡号",
    "交易金额",
    "清算金额",
    "手续费",
    "T0手续费",
    "D1手续费",
    "流水号",
    "检索号",
    "卡类型",
    "发卡行",
    "商户号",
    "商户名称",
    "分店简称",
    "商户订单号",
    "银商订单号",
    "交易方式",
    "分店",
    "优惠金额",
    "分期手续费",
    "付款附言",
    "备注",
    "买家ID",
];

const COLUMN_TYPES: [ColumnType; 26] = [
    ColumnType::Date,
    ColumnType::DateTime,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Decimal(DecimalScale::Two),
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
];

const MERCHANTS: [&str; 2] = ["89813014812B06R", "89813015722APT1"];

pub(crate) const D1_NOTICE: &str =
    "请注意：D1手续费字段为预估数据仅供参考，实际以17：40分之后的D1划付数据为准。";

/// 一条有效的门店银联交易明细（已排除首行汇总、表头及末行提示）。
#[derive(Debug)]
pub(crate) struct UnionPayRecord {
    pub settlement_date: Value,
    pub transaction_time: Value,
    pub terminal_no: String,
    pub transaction_type: String,
    pub card_no: String,
    pub transaction_amount: Option<Decimal>,
    pub settlement_amount: Option<Decimal>,
    pub fee: Option<Decimal>,
    pub t0_fee: Option<Decimal>,
    pub d1_fee: Option<Decimal>,
    pub serial_no: String,
    pub retrieval_no: String,
    pub card_type: String,
    pub issuing_bank: String,
    pub merchant_no: String,
    pub merchant_name: String,
    pub branch_short_name: String,
    pub merchant_order_no: String,
    pub acquirer_order_no: String,
    pub transaction_method: String,
    pub branch: String,
    pub discount_amount: Option<Decimal>,
    pub installment_fee: Option<Decimal>,
    pub payment_note: String,
    pub remark: String,
    pub buyer_id: String,
}

/// 文件名须完整符合`商户号_MX_YYYYMMDDHHMMSS_序号.xlsx`，商户号属于第 6.1 节范围。
fn matches_filename(name: &str) -> bool {
    let Some(stem) = name.strip_suffix(".xlsx") else {
        return false;
    };
    let Some(merchant) = MERCHANTS.iter().find(|m| stem.starts_with(**m)) else {
        return false;
    };
    let Some(rest) = stem[merchant.len()..].strip_prefix("_MX_") else {
        return false;
    };
    match rest.split_once('_') {
        Some((timestamp, seq)) => {
            timestamp.len() == 14
                && timestamp.bytes().all(|b| b.is_ascii_digit())
                && !seq.is_empty()
        }
        None => false,
    }
}

fn read_record(
    sheet: &SheetGrid,
    row: u32,
    file: &str,
    sheet_name: &str,
) -> Result<UnionPayRecord, ProcessError> {
    let cells: Vec<RawCell> = (1..=26).map(|col| sheet.cell(row, col)).collect();

    let text_at = |index: usize| -> Result<String, ProcessError> {
        cell_text(&cells[index - 1]).map_err(|detail| {
            data_error(
                file,
                sheet_name,
                row,
                HEADERS[index - 1],
                cells[index - 1].to_string(),
                detail,
            )
        })
    };
    let amount_at = |index: usize| -> Result<Option<Decimal>, ProcessError> {
        cell_amount(&cells[index - 1]).map_err(|detail| {
            data_error(
                file,
                sheet_name,
                row,
                HEADERS[index - 1],
                cells[index - 1].to_string(),
                detail,
            )
        })
    };

    Ok(UnionPayRecord {
        settlement_date: cell_date_or_text(&cells[0]),
        transaction_time: cell_datetime_or_text(&cells[1]),
        terminal_no: text_at(3)?,
        transaction_type: text_at(4)?,
        card_no: text_at(5)?,
        transaction_amount: amount_at(6)?,
        settlement_amount: amount_at(7)?,
        fee: amount_at(8)?,
        t0_fee: amount_at(9)?,
        d1_fee: amount_at(10)?,
        serial_no: text_at(11)?,
        retrieval_no: text_at(12)?,
        card_type: text_at(13)?,
        issuing_bank: text_at(14)?,
        merchant_no: text_at(15)?,
        merchant_name: text_at(16)?,
        branch_short_name: text_at(17)?,
        merchant_order_no: text_at(18)?,
        acquirer_order_no: text_at(19)?,
        transaction_method: text_at(20)?,
        branch: text_at(21)?,
        discount_amount: amount_at(22)?,
        installment_fee: amount_at(23)?,
        payment_note: text_at(24)?,
        remark: text_at(25)?,
        buyer_id: text_at(26)?,
    })
}

/// 文件发现 → 工作表与表头校验 → 排除首行汇总、表头、末行提示 → 重复导出检查
/// → 返回有效原始交易。`coupons.rs`复用本函数建立参考号校验集。
pub(crate) fn load_records(input_dir: &Path) -> Result<Vec<UnionPayRecord>, ProcessError> {
    let mut files: Vec<_> = list_xlsx_files(input_dir)?
        .into_iter()
        .filter(|path| {
            path.file_name()
                .and_then(|n| n.to_str())
                .is_some_and(matches_filename)
        })
        .collect();
    files.sort();
    if files.is_empty() {
        return Err(ProcessError::NoInput {
            pattern: "商户号_MX_YYYYMMDDHHMMSS_序号.xlsx".to_string(),
        });
    }

    let mut records = Vec::new();
    let mut fingerprints: HashMap<String, String> = HashMap::new();

    for path in &files {
        let file_name = path.file_name().unwrap().to_string_lossy().into_owned();
        let sheets = open_sheets(path)?;

        for sheet in &sheets {
            let sheet_name = sheet.name().to_string();

            let header = sheet.row_texts(2);
            if header.iter().map(String::as_str).collect::<Vec<_>>() != HEADERS {
                return Err(ProcessError::Structure {
                    file: file_name,
                    sheet: sheet_name,
                    detail: format!("第2行表头与规定的26个字段不一致：{header:?}"),
                });
            }

            let last_row = sheet.last_value_row().unwrap_or(0);
            if last_row < 3 {
                return Err(ProcessError::Structure {
                    file: file_name,
                    sheet: sheet_name,
                    detail: "缺少有效交易数据区域（第3行至倒数第2行）".to_string(),
                });
            }
            let notice = sheet.cell(last_row, 1).to_string();
            if notice != D1_NOTICE {
                return Err(ProcessError::Structure {
                    file: file_name,
                    sheet: sheet_name,
                    detail: format!("末行提示文字与规定不符：{notice:?}"),
                });
            }

            let fingerprint = sheet.fingerprint(1, last_row.saturating_sub(1));
            check_duplicate_fingerprint(&mut fingerprints, fingerprint, &file_name)?;

            for row in 3..last_row {
                records.push(read_record(sheet, row, &file_name, &sheet_name)?);
            }
        }
    }

    Ok(records)
}

fn is_return(transaction_type: &str) -> bool {
    matches!(transaction_type, "消费撤销" | "消费撤消" | "联机退货")
}

/// 退单按非空`检索号`回溯唯一的原`消费`记录；无命中或同号多条消费时不任选。
fn returned_original_refs(records: &[UnionPayRecord]) -> HashSet<String> {
    let mut consumption_counts: HashMap<&str, usize> = HashMap::new();
    for record in records
        .iter()
        .filter(|r| r.transaction_type == "消费" && !r.retrieval_no.is_empty())
    {
        *consumption_counts.entry(&record.retrieval_no).or_default() += 1;
    }

    records
        .iter()
        .filter(|r| is_return(&r.transaction_type))
        .filter(|r| {
            !r.retrieval_no.is_empty()
                && consumption_counts.get(r.retrieval_no.as_str()) == Some(&1)
        })
        .map(|r| r.retrieval_no.clone())
        .collect()
}

fn to_row(record: UnionPayRecord, fill: Option<Fill>) -> Row {
    let values = vec![
        record.settlement_date,
        record.transaction_time,
        text_value(record.terminal_no),
        text_value(record.transaction_type),
        text_value(record.card_no),
        amount_value(record.transaction_amount),
        amount_value(record.settlement_amount),
        amount_value(record.fee),
        amount_value(record.t0_fee),
        amount_value(record.d1_fee),
        text_value(record.serial_no),
        text_value(record.retrieval_no),
        text_value(record.card_type),
        text_value(record.issuing_bank),
        text_value(record.merchant_no),
        text_value(record.merchant_name),
        text_value(record.branch_short_name),
        text_value(record.merchant_order_no),
        text_value(record.acquirer_order_no),
        text_value(record.transaction_method),
        text_value(record.branch),
        amount_value(record.discount_amount),
        amount_value(record.installment_fee),
        text_value(record.payment_note),
        text_value(record.remark),
        text_value(record.buyer_id),
    ];
    Row { values, fill }
}

fn columns() -> Vec<Column> {
    HEADERS
        .iter()
        .zip(COLUMN_TYPES)
        .map(|(&name, ty)| Column { name, ty })
        .collect()
}

pub struct UnionPayJob;

impl Job for UnionPayJob {
    fn category(&self) -> Category {
        Category::UnionPay
    }

    fn title(&self) -> &'static str {
        "门店银联交易明细"
    }

    fn output_stem(&self) -> &'static str {
        "银联交易明细门店"
    }

    fn run(&self, input_dir: &Path) -> Result<Table, ProcessError> {
        let records = load_records(input_dir)?;
        let returned_original_refs = returned_original_refs(&records);

        // 稳定分区：退单及其唯一原消费一并沉底，两个区域各自保持原相对顺序。
        let (normal, returned): (Vec<_>, Vec<_>) = records.into_iter().partition(|r| {
            !is_return(&r.transaction_type)
                && !(r.transaction_type == "消费"
                    && returned_original_refs.contains(&r.retrieval_no))
        });

        let rows = normal
            .into_iter()
            .map(|record| to_row(record, None))
            .chain(returned.into_iter().map(|mut record| {
                record.remark = "已退货".to_string();
                to_row(record, Some(Fill::Pink))
            }))
            .collect();

        Ok(Table {
            columns: columns(),
            rows,
        })
    }
}

#[cfg(test)]
mod tests {
    use rust_xlsxwriter::{Workbook, Worksheet};

    use super::*;
    use crate::test_support::unique_temp_path;

    // 交易金额、清算金额、手续费、T0手续费、D1手续费、优惠金额、分期手续费的零基列号。
    const AMOUNT_COLS: [usize; 7] = [5, 6, 7, 8, 9, 21, 22];

    fn write_row(sheet: &mut Worksheet, row: u32, values: &[&str; 26]) {
        for (col, value) in values.iter().enumerate() {
            if AMOUNT_COLS.contains(&col) {
                if !value.is_empty() {
                    sheet
                        .write_number(row, col as u16, value.parse::<f64>().unwrap())
                        .unwrap();
                }
            } else {
                sheet.write_string(row, col as u16, *value).unwrap();
            }
        }
    }

    fn write_workbook(path: &Path, summary: &str, rows: &[[&str; 26]]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("对账数据").unwrap();
        sheet.write_string(0, 0, summary).unwrap();
        for (col, header) in HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        for (index, row) in rows.iter().enumerate() {
            write_row(sheet, (2 + index) as u32, row);
        }
        sheet
            .write_string(2 + rows.len() as u32, 0, D1_NOTICE)
            .unwrap();
        workbook.save(path).unwrap();
    }

    /// 一条示例交易：清算时间`col0`、交易时间`col1`为文本，其余字段按参数覆盖。
    fn sample_row<'a>(
        order_no: &'a str,
        transaction_type: &'a str,
        remark: &'a str,
        retrieval_no: &'a str,
    ) -> [&'a str; 26] {
        [
            "20260914",
            "2026-09-14 10:18:09",
            "T001",
            transaction_type,
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
            order_no,
            "AC0001",
            "云闪付",
            "分店A",
            "0.00",
            "0.00",
            "备注文字",
            remark,
            "buyer001",
        ]
    }

    #[test]
    fn ignores_files_with_unmatched_names() {
        assert!(matches_filename("89813014812B06R_MX_20260914101809_1.xlsx"));
        assert!(!matches_filename("89813014812B06R_MX_20260914101809_1.csv"));
        assert!(!matches_filename("其他商户_MX_20260914101809_1.xlsx"));
        assert!(!matches_filename(
            "~$89813014812B06R_MX_20260914101809_1.xlsx"
        ));
        assert!(!matches_filename("89813014812B06R_MX_2026_1.xlsx"));
    }

    #[test]
    fn merges_files_classifies_returns_and_moves_them_to_bottom() {
        let dir = unique_temp_path("unionpay-happy-path");
        std::fs::create_dir_all(&dir).unwrap();

        write_workbook(
            &dir.join("89813014812B06R_MX_20260914101809_1.xlsx"),
            "汇总1",
            &[
                sample_row("ORDER-A", "消费", "", "16867252734N"),
                sample_row("ORDER-B", "消费撤消", "原备注", "16867252734N"),
            ],
        );
        write_workbook(
            &dir.join("89813015722APT1_MX_20260914101900_1.xlsx"),
            "汇总2",
            &[sample_row("ORDER-C", "消费", "", "16867252735N")],
        );
        // 不符合命名规则的文件必须被忽略。
        std::fs::write(dir.join("说明.xlsx"), b"not a real workbook").unwrap();

        let job = UnionPayJob;
        let table = job.run(&dir).unwrap();

        assert_eq!(table.columns.len(), 26);
        assert_eq!(table.rows.len(), 3);

        let order_no_col = 17; // 商户订单号
        let order_nos: Vec<_> = table
            .rows
            .iter()
            .map(|row| match &row.values[order_no_col] {
                Value::Text(text) => text.as_str(),
                _ => panic!("expected text"),
            })
            .collect();
        // 未退消费在前；被退原消费和退单沉底，沉底区域保持原相对顺序。
        assert_eq!(order_nos, vec!["ORDER-C", "ORDER-A", "ORDER-B"]);

        for returned_row in &table.rows[1..] {
            assert_eq!(returned_row.fill, Some(Fill::Pink));
            assert_eq!(returned_row.values[24], Value::Text("已退货".to_string()));
        }
        assert_eq!(table.rows[0].fill, None);

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn ambiguous_original_consumptions_are_not_marked_as_returned() {
        let dir = unique_temp_path("unionpay-ambiguous-original");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("89813014812B06R_MX_20260914101809_1.xlsx"),
            "汇总",
            &[
                sample_row("ORDER-A", "消费", "原备注A", "16867252734N"),
                sample_row("ORDER-B", "消费", "原备注B", "16867252734N"),
                sample_row("ORDER-C", "联机退货", "退单原备注", "16867252734N"),
            ],
        );

        let table = UnionPayJob.run(&dir).unwrap();

        assert_eq!(table.rows[0].fill, None);
        assert_eq!(table.rows[0].values[24], Value::Text("原备注A".to_string()));
        assert_eq!(table.rows[1].fill, None);
        assert_eq!(table.rows[1].values[24], Value::Text("原备注B".to_string()));
        assert_eq!(table.rows[2].fill, Some(Fill::Pink));
        assert_eq!(table.rows[2].values[24], Value::Text("已退货".to_string()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_header_mismatch() {
        let dir = unique_temp_path("unionpay-bad-header");
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("89813014812B06R_MX_20260914101809_1.xlsx");

        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("对账数据").unwrap();
        sheet.write_string(0, 0, "汇总").unwrap();
        sheet.write_string(1, 0, "错误表头").unwrap();
        sheet.write_string(2, 0, D1_NOTICE).unwrap();
        workbook.save(&path).unwrap();

        let error = load_records(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Structure { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_nothing_matches() {
        let dir = unique_temp_path("unionpay-no-input");
        std::fs::create_dir_all(&dir).unwrap();

        let error = load_records(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_wrong_trailing_notice() {
        let dir = unique_temp_path("unionpay-bad-notice");
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("89813014812B06R_MX_20260914101809_1.xlsx");
        write_workbook(
            &path,
            "汇总",
            &[sample_row("ORDER-A", "消费", "", "16867252734N")],
        );

        // 破坏末行提示文字。
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("对账数据").unwrap();
        sheet.write_string(0, 0, "汇总").unwrap();
        for (col, header) in HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        write_row(sheet, 2, &sample_row("ORDER-A", "消费", "", "16867252734N"));
        sheet.write_string(3, 0, "不是规定的提示文字").unwrap();
        workbook.save(&path).unwrap();

        let error = load_records(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Structure { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_duplicate_export_across_files() {
        let dir = unique_temp_path("unionpay-duplicate");
        std::fs::create_dir_all(&dir).unwrap();

        let rows = [sample_row("ORDER-A", "消费", "", "16867252734N")];
        write_workbook(
            &dir.join("89813014812B06R_MX_20260914101809_1.xlsx"),
            "汇总",
            &rows,
        );
        write_workbook(
            &dir.join("89813014812B06R_MX_20260914101809_2.xlsx"),
            "汇总",
            &rows,
        );

        let error = load_records(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Duplicate { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }
}
