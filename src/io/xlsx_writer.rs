use std::path::Path;

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Workbook};

use crate::model::{Column, ColumnType, DecimalScale, Fill, ProcessError, Table, Value};

fn base_format(ty: ColumnType) -> Format {
    let format = Format::new()
        .set_font_name("微软雅黑")
        .set_font_size(11)
        .set_align(FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin)
        .set_border_color(Color::RGB(0xD9D9D9));

    match ty {
        ColumnType::Text => format.set_align(FormatAlign::Left),
        ColumnType::Decimal(DecimalScale::Original) => format.set_align(FormatAlign::Right),
        ColumnType::Decimal(DecimalScale::Two) => {
            format.set_num_format("0.00").set_align(FormatAlign::Right)
        }
        ColumnType::Integer => format.set_num_format("0").set_align(FormatAlign::Right),
        ColumnType::Ratio => format.set_num_format("0.00%").set_align(FormatAlign::Right),
        ColumnType::Date => format
            .set_num_format("yyyy-mm-dd")
            .set_align(FormatAlign::Center),
        ColumnType::Time => format
            .set_num_format("hh:mm:ss")
            .set_align(FormatAlign::Center),
        ColumnType::DateTime => format
            .set_num_format("yyyy-mm-dd hh:mm:ss")
            .set_align(FormatAlign::Center),
    }
}

/// 按列类型确定数字格式，再按行填色叠加背景色；`rust_xlsxwriter`会自动去重相同的`Format`。
fn cell_format(column: &Column, fill: Option<Fill>) -> Format {
    let format = base_format(column.ty);
    match fill {
        Some(Fill::Yellow) => format.set_background_color(Color::RGB(0xFFEB9C)),
        Some(Fill::Pink) => format.set_background_color(Color::RGB(0xFFC7CE)),
        None => format,
    }
}

/// 把`Table`写成单工作表 XLSX 文件；调用方负责临时文件命名与正式替换。
pub fn write_table(table: &Table, path: &Path) -> Result<(), ProcessError> {
    table.validate()?;

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let header_format = Format::new()
        .set_font_name("微软雅黑")
        .set_font_size(11)
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color(Color::RGB(0x5B9BD5))
        .set_align(FormatAlign::Left)
        .set_align(FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin)
        .set_border_color(Color::RGB(0xD9D9D9));

    worksheet.set_default_row_height(22);
    worksheet
        .set_row_height(0, 30)
        .map_err(|error| ProcessError::Io(std::io::Error::other(error)))?;
    worksheet
        .set_freeze_panes(1, 0)
        .map_err(|error| ProcessError::Io(std::io::Error::other(error)))?;

    for (col_index, column) in table.columns.iter().enumerate() {
        worksheet
            .write_string_with_format(0, col_index as u16, column.name, &header_format)
            .map_err(|error| ProcessError::Io(std::io::Error::other(error)))?;
    }

    for (row_index, row) in table.rows.iter().enumerate() {
        let excel_row = (row_index + 1) as u32;
        for (col_index, value) in row.values.iter().enumerate() {
            let column = &table.columns[col_index];
            let col = col_index as u16;
            let format = cell_format(column, row.fill);
            let result = match value {
                Value::Empty => worksheet.write_blank(excel_row, col, &format).map(|_| ()),
                Value::Text(text) => worksheet
                    .write_with_format(excel_row, col, text, &format)
                    .map(|_| ()),
                Value::Decimal(amount) => worksheet
                    .write_with_format(excel_row, col, *amount, &format)
                    .map(|_| ()),
                Value::Integer(number) => worksheet
                    .write_with_format(excel_row, col, *number, &format)
                    .map(|_| ()),
                Value::Ratio(ratio) => worksheet
                    .write_with_format(excel_row, col, *ratio, &format)
                    .map(|_| ()),
                Value::Date(date) => worksheet
                    .write_with_format(excel_row, col, date, &format)
                    .map(|_| ()),
                Value::Time(time) => worksheet
                    .write_with_format(excel_row, col, time, &format)
                    .map(|_| ()),
                Value::DateTime(datetime) => worksheet
                    .write_with_format(excel_row, col, datetime, &format)
                    .map(|_| ()),
            };
            result.map_err(|error| ProcessError::Io(std::io::Error::other(error)))?;
        }
    }

    worksheet
        .set_autofit_max_row(200)
        .set_autofit_max_width(300)
        .autofit();

    let widths: &[f64] = match table.columns.as_slice() {
        columns
            if columns.len() == 26
                && columns[0].name == "清算时间"
                && columns[25].name == "买家ID" =>
        {
            &[
                20.0, 20.0, 14.0, 12.0, 22.0, 14.0, 14.0, 12.0, 12.0, 12.0, 16.0, 18.0, 12.0, 16.0,
                18.0, 24.0, 16.0, 28.0, 28.0, 14.0, 16.0, 14.0, 14.0, 14.0, 16.0, 18.0,
            ]
        }
        columns
            if columns.len() == 24
                && columns[0].name == "拨付批次"
                && columns[23].name == "原拨付批次" =>
        {
            &[
                32.0, 20.0, 16.0, 26.0, 26.0, 22.0, 18.0, 14.0, 14.0, 14.0, 14.0, 12.0, 24.0, 14.0,
                16.0, 12.0, 16.0, 36.0, 14.0, 24.0, 18.0, 24.0, 48.0, 18.0,
            ]
        }
        _ => &[],
    };
    for (col, width) in widths.iter().enumerate() {
        worksheet
            .set_column_width(col as u16, *width)
            .map_err(|error| ProcessError::Io(std::io::Error::other(error)))?;
    }

    workbook
        .save(path)
        .map_err(|error| ProcessError::Io(std::io::Error::other(error)))
}

#[cfg(test)]
mod tests {
    use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
    use rust_decimal::Decimal;
    use rust_decimal::prelude::FromPrimitive;

    use super::*;
    use crate::io::xlsx_reader::{RawCell, open_sheets};
    use crate::model::{Row, Table};
    use crate::test_support::unique_temp_path;

    fn column(name: &'static str, ty: ColumnType) -> Column {
        Column { name, ty }
    }

    /// 验证写入后再用`io::xlsx_reader`读回，各类型的值都能正确往返，
    /// 且行填色不影响单元格数值。
    #[test]
    fn round_trips_every_value_kind() {
        let columns = vec![
            column("文本", ColumnType::Text),
            column("金额原精度", ColumnType::Decimal(DecimalScale::Original)),
            column("金额两位小数", ColumnType::Decimal(DecimalScale::Two)),
            column("整数", ColumnType::Integer),
            column("比例", ColumnType::Ratio),
            column("日期", ColumnType::Date),
            column("时间", ColumnType::Time),
            column("日期时间", ColumnType::DateTime),
            column("空值", ColumnType::Text),
        ];

        let row = Row {
            values: vec![
                Value::Text("带,逗号".to_string()),
                Value::Decimal(Decimal::from_f64(1234.5678).unwrap()),
                Value::Decimal(Decimal::from_f64(99.5).unwrap()),
                Value::Integer(42),
                Value::Ratio(Decimal::from_f64(0.15).unwrap()),
                Value::Date(NaiveDate::from_ymd_opt(2026, 7, 28).unwrap()),
                Value::Time(NaiveTime::from_hms_opt(13, 5, 9).unwrap()),
                Value::DateTime(NaiveDateTime::new(
                    NaiveDate::from_ymd_opt(2026, 9, 14).unwrap(),
                    NaiveTime::from_hms_opt(10, 18, 9).unwrap(),
                )),
                Value::Empty,
            ],
            fill: Some(Fill::Pink),
        };

        let table = Table {
            columns,
            rows: vec![row],
        };
        let path = unique_temp_path("xlsx-writer-round-trip").with_extension("xlsx");

        write_table(&table, &path).unwrap();

        let sheets = open_sheets(&path).unwrap();
        assert_eq!(sheets.len(), 1);
        let sheet = &sheets[0];

        assert_eq!(
            sheet.row_texts(1),
            vec![
                "文本",
                "金额原精度",
                "金额两位小数",
                "整数",
                "比例",
                "日期",
                "时间",
                "日期时间",
                "空值",
            ]
        );

        assert_eq!(sheet.cell(2, 1), RawCell::Text("带,逗号".to_string()));
        assert!(matches!(sheet.cell(2, 2), RawCell::Float(v) if (v - 1234.5678).abs() < 1e-9));
        assert!(matches!(sheet.cell(2, 3), RawCell::Float(v) if (v - 99.5).abs() < 1e-9));
        // rust_xlsxwriter 把 i64 写成普通数值单元格，calamine 读回后一律归类为 Float。
        assert!(matches!(sheet.cell(2, 4), RawCell::Float(v) if (v - 42.0).abs() < 1e-9));
        assert!(matches!(sheet.cell(2, 5), RawCell::Float(v) if (v - 0.15).abs() < 1e-9));
        assert!(matches!(sheet.cell(2, 6), RawCell::DateTime(_)));
        assert!(matches!(sheet.cell(2, 7), RawCell::DateTime(_)));
        assert!(matches!(sheet.cell(2, 8), RawCell::DateTime(_)));
        assert_eq!(sheet.cell(2, 9), RawCell::Empty);

        std::fs::remove_file(&path).ok();
    }
}
