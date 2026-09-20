use std::collections::{HashMap, HashSet};
use std::path::Path;

use crate::io::paths::list_xlsx_files;
use crate::io::xlsx_reader::{SheetGrid, open_sheets};
use crate::model::{Column, ColumnType, DecimalScale, ProcessError, Row, Table, Value};
use crate::utils::natural_sort;

use super::{
    Category, Job, amount_value, cell_amount, cell_text, check_duplicate_fingerprint, data_error,
    parse_date_field, parse_time_field, text_value,
};

const HEADERS: [&str; 33] = [
    "银商订单号",
    "银商原订单号",
    "交易类型",
    "交易日期",
    "交易时间",
    "支付方式",
    "订单金额",
    "应收金额",
    "商品名称",
    "商品明细",
    "品牌",
    "商品类别",
    "能耗等级",
    "商品型号",
    "商品识别码SN",
    "IMEI1",
    "IMEI2",
    "备注",
    "交易参考号",
    "商家优惠金额",
    "渠道优惠（云闪付、宝信）",
    "活动ID",
    "实收金额",
    "银商商户号",
    "银商商户名称",
    "终端号",
    "门店编码",
    "门店名称",
    "小U交易流水号",
    "小U原交易流水号",
    "订单号",
    "订单流水号",
    "收银员",
];

const COLUMN_TYPES: [ColumnType; 33] = [
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Text,
    ColumnType::Date,
    ColumnType::Time,
    ColumnType::Text,
    ColumnType::Decimal(DecimalScale::Original),
    ColumnType::Decimal(DecimalScale::Original),
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
    ColumnType::Decimal(DecimalScale::Original),
    ColumnType::Decimal(DecimalScale::Original),
    ColumnType::Text,
    ColumnType::Decimal(DecimalScale::Original),
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
];

// 列位置（1 基）：小U交易流水号、订单号，用于合并后的全局重复检查。
const XIAOU_SERIAL_COL: usize = 29;
const ORDER_NO_COL: usize = 31;

fn matches_filename(name: &str) -> bool {
    name.contains("银联国补明细")
}

fn standardize_brand(value: String) -> String {
    match value.as_str() {
        "Leader" | "leader" => "统帅".to_string(),
        "A.O.史密斯" => "AO史密斯".to_string(),
        _ => value,
    }
}

fn standardize_category(value: String) -> String {
    match value.as_str() {
        "热水器" => "厨卫".to_string(),
        "电视机" => "电视".to_string(),
        "电冰箱" => "冰箱".to_string(),
        _ => value,
    }
}

fn read_row(
    sheet: &SheetGrid,
    row: u32,
    file: &str,
    sheet_name: &str,
) -> Result<Row, ProcessError> {
    let text_at = |col: u32, field: &'static str| -> Result<String, ProcessError> {
        let cell = sheet.cell(row, col);
        cell_text(&cell)
            .map_err(|detail| data_error(file, sheet_name, row, field, cell.to_string(), detail))
    };
    let amount_at = |col: u32, field: &'static str| -> Result<Value, ProcessError> {
        let cell = sheet.cell(row, col);
        cell_amount(&cell)
            .map(amount_value)
            .map_err(|detail| data_error(file, sheet_name, row, field, cell.to_string(), detail))
    };

    let values = vec![
        text_value(text_at(1, "银商订单号")?),
        text_value(text_at(2, "银商原订单号")?),
        text_value(text_at(3, "交易类型")?),
        parse_date_field(&sheet.cell(row, 4), "交易日期", file, sheet_name, row)?,
        parse_time_field(&sheet.cell(row, 5), "交易时间", file, sheet_name, row)?,
        text_value(text_at(6, "支付方式")?),
        amount_at(7, "订单金额")?,
        amount_at(8, "应收金额")?,
        text_value(text_at(9, "商品名称")?),
        text_value(text_at(10, "商品明细")?),
        text_value(standardize_brand(text_at(11, "品牌")?)),
        text_value(standardize_category(text_at(12, "商品类别")?)),
        text_value(text_at(13, "能耗等级")?),
        text_value(text_at(14, "商品型号")?),
        text_value(text_at(15, "商品识别码SN")?),
        text_value(text_at(16, "IMEI1")?),
        text_value(text_at(17, "IMEI2")?),
        text_value(text_at(18, "备注")?),
        text_value(text_at(19, "交易参考号")?),
        amount_at(20, "商家优惠金额")?,
        amount_at(21, "渠道优惠（云闪付、宝信）")?,
        text_value(text_at(22, "活动ID")?),
        amount_at(23, "实收金额")?,
        text_value(text_at(24, "银商商户号")?),
        text_value(text_at(25, "银商商户名称")?),
        text_value(text_at(26, "终端号")?),
        text_value(text_at(27, "门店编码")?),
        text_value(text_at(28, "门店名称")?),
        text_value(text_at(29, "小U交易流水号")?),
        text_value(text_at(30, "小U原交易流水号")?),
        text_value(text_at(31, "订单号")?),
        text_value(text_at(32, "订单流水号")?),
        text_value(text_at(33, "收银员")?),
    ];
    Ok(Row { values, fill: None })
}

pub struct UnionSubsidyJob;

impl Job for UnionSubsidyJob {
    fn category(&self) -> Category {
        Category::UnionSubsidy
    }

    fn title(&self) -> &'static str {
        "银联国补明细"
    }

    fn output_stem(&self) -> &'static str {
        "银联交易明细所有"
    }

    fn run(&self, input_dir: &Path) -> Result<Table, ProcessError> {
        let mut files: Vec<_> = list_xlsx_files(input_dir)?
            .into_iter()
            .filter(|path| {
                path.file_name()
                    .and_then(|n| n.to_str())
                    .is_some_and(matches_filename)
            })
            .collect();
        files.sort_by(|a, b| {
            natural_sort::natural_cmp(
                &a.file_name().unwrap().to_string_lossy(),
                &b.file_name().unwrap().to_string_lossy(),
            )
        });
        if files.is_empty() {
            return Err(ProcessError::NoInput {
                pattern: "*银联国补明细*.xlsx".to_string(),
            });
        }

        let mut rows = Vec::new();
        let mut fingerprints: HashMap<String, String> = HashMap::new();
        let mut seen_xiaou_serial: HashSet<String> = HashSet::new();
        let mut seen_order_no: HashSet<String> = HashSet::new();

        for path in &files {
            let file_name = path.file_name().unwrap().to_string_lossy().into_owned();
            let sheets = open_sheets(path)?;

            for sheet in &sheets {
                let sheet_name = sheet.name().to_string();

                let header = sheet.row_texts(1);
                if header.iter().map(String::as_str).collect::<Vec<_>>() != HEADERS {
                    return Err(ProcessError::Structure {
                        file: file_name,
                        sheet: sheet_name,
                        detail: format!("第1行表头与规定的33个字段不一致：{header:?}"),
                    });
                }

                let last_row = sheet.last_value_row().unwrap_or(1);

                let fingerprint = sheet.fingerprint(1, last_row);
                check_duplicate_fingerprint(&mut fingerprints, fingerprint, &file_name)?;

                for row in 2..=last_row {
                    let record_row = read_row(sheet, row, &file_name, &sheet_name)?;

                    if let Value::Text(no) = &record_row.values[XIAOU_SERIAL_COL - 1]
                        && !seen_xiaou_serial.insert(no.clone())
                    {
                        return Err(ProcessError::Duplicate {
                            detail: format!("小U交易流水号重复：{no}"),
                        });
                    }
                    if let Value::Text(no) = &record_row.values[ORDER_NO_COL - 1]
                        && !seen_order_no.insert(no.clone())
                    {
                        return Err(ProcessError::Duplicate {
                            detail: format!("订单号重复：{no}"),
                        });
                    }

                    rows.push(record_row);
                }
            }
        }

        Ok(Table {
            columns: output_columns(),
            rows,
        })
    }
}

fn output_columns() -> Vec<Column> {
    HEADERS
        .iter()
        .zip(COLUMN_TYPES)
        .map(|(&name, ty)| Column { name, ty })
        .collect()
}

#[cfg(test)]
mod tests {
    use rust_xlsxwriter::Workbook;

    use super::*;
    use crate::test_support::unique_temp_path;

    fn sample_row<'a>(order_no: &'a str, brand: &'a str, category: &'a str) -> [&'a str; 33] {
        [
            order_no,
            "ORIG0001",
            "国补",
            "2026-09-14",
            "10:18:09",
            "银联",
            "1000.00",
            "900.00",
            "手机",
            "手机明细",
            brand,
            category,
            "一级",
            "型号A",
            "SN0001",
            "IMEI001",
            "IMEI002",
            "备注",
            "16867252734N",
            "50.00",
            "10.00",
            "ACT001",
            "890.00",
            "MCH001",
            "某商户",
            "T001",
            "S001",
            "某门店",
            order_no, // 小U交易流水号：用 order_no 保证唯一，便于测试断言
            "XORIG0001",
            order_no,
            "F0001",
            "收银员甲",
        ]
    }

    fn write_workbook(path: &Path, rows: &[[&str; 33]]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        for (col, header) in HEADERS.iter().enumerate() {
            sheet.write_string(0, col as u16, *header).unwrap();
        }
        for (row_index, row) in rows.iter().enumerate() {
            for (col, value) in row.iter().enumerate() {
                sheet
                    .write_string((1 + row_index) as u32, col as u16, *value)
                    .unwrap();
            }
        }
        workbook.save(path).unwrap();
    }

    #[test]
    fn ignores_files_without_marker_text() {
        assert!(matches_filename("银联国补明细1.xlsx"));
        assert!(matches_filename("前缀银联国补明细.xlsx"));
        assert!(!matches_filename("发票_20260914.xlsx"));
    }

    #[test]
    fn merges_in_natural_filename_order_and_standardizes_values() {
        let dir = unique_temp_path("union-subsidy-happy-path");
        std::fs::create_dir_all(&dir).unwrap();

        // 字符串顺序会把 "10" 排在 "2" 之前，自然排序必须反过来。
        write_workbook(
            &dir.join("银联国补明细10.xlsx"),
            &[sample_row("ORDER-FILE10", "Leader", "热水器")],
        );
        write_workbook(
            &dir.join("银联国补明细2.xlsx"),
            &[sample_row("ORDER-FILE2", "leader", "电视机")],
        );
        // 不含标记文字的文件必须被忽略。
        std::fs::write(dir.join("说明.xlsx"), b"not a real workbook").unwrap();

        let table = UnionSubsidyJob.run(&dir).unwrap();

        assert_eq!(table.columns.len(), 33);
        assert_eq!(table.rows.len(), 2);

        let order_no = |row: &Row| match &row.values[0] {
            Value::Text(text) => text.clone(),
            _ => panic!("expected text"),
        };
        assert_eq!(order_no(&table.rows[0]), "ORDER-FILE2");
        assert_eq!(order_no(&table.rows[1]), "ORDER-FILE10");

        // 品牌标准化：Leader/leader 均替换为“统帅”。
        assert_eq!(table.rows[0].values[10], Value::Text("统帅".to_string()));
        assert_eq!(table.rows[1].values[10], Value::Text("统帅".to_string()));
        // 商品类别标准化：电视机 → 电视，热水器 → 厨卫。
        assert_eq!(table.rows[0].values[11], Value::Text("电视".to_string()));
        assert_eq!(table.rows[1].values[11], Value::Text("厨卫".to_string()));

        // 交易日期/交易时间被正确解析为 Date/Time。
        assert!(matches!(table.rows[0].values[3], Value::Date(_)));
        assert!(matches!(table.rows[0].values[4], Value::Time(_)));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_header_mismatch() {
        let dir = unique_temp_path("union-subsidy-bad-header");
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("银联国补明细1.xlsx");

        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, "错误表头").unwrap();
        sheet.write_string(1, 0, "数据").unwrap();
        workbook.save(&path).unwrap();

        let error = UnionSubsidyJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Structure { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_duplicate_xiaou_serial_across_files() {
        let dir = unique_temp_path("union-subsidy-dup-xiaou");
        std::fs::create_dir_all(&dir).unwrap();

        // 除小U交易流水号/订单号（沿用同一个 order_no）外，其余字段（品牌）不同，
        // 确保命中的是逐字段重复检查，而不是整表内容指纹重复检查。
        write_workbook(
            &dir.join("银联国补明细1.xlsx"),
            &[sample_row("ORDER-A", "格力", "空调")],
        );
        write_workbook(
            &dir.join("银联国补明细2.xlsx"),
            &[sample_row("ORDER-A", "美的", "空调")],
        );

        let error = UnionSubsidyJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Duplicate { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_whole_sheet_duplicate_export_across_files() {
        let dir = unique_temp_path("union-subsidy-dup-sheet");
        std::fs::create_dir_all(&dir).unwrap();

        // 两个不同文件、完全相同的表头+明细内容（包括各字段值），构成整表重复导出。
        let rows = [sample_row("ORDER-SAME", "格力", "空调")];
        write_workbook(&dir.join("银联国补明细A.xlsx"), &rows);
        write_workbook(&dir.join("银联国补明细B.xlsx"), &rows);

        let error = UnionSubsidyJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Duplicate { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_nothing_matches() {
        let dir = unique_temp_path("union-subsidy-no-input");
        std::fs::create_dir_all(&dir).unwrap();

        let error = UnionSubsidyJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }
}
