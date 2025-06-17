# dealwithmileagexlsx

This repository contains Excel sheets describing communication line crossings and migrations.

Use `parse_mileage.py` to parse mileage information in these spreadsheets.

```
python parse_mileage.py <xlsx_path> <column_letter>
```

- `xlsx_path`: path to the Excel file.
- `column_letter`: the column containing mileage values (e.g. `B` or `C`).

The script replaces the mileage column with **起始里程** and adds a new **结束里程** column.
The processed workbook is saved as `<original_name>_里程解析后.xlsx`.

To align two already-parsed spreadsheets by mileage, use `align_mileage.py`:

```
python align_mileage.py <table1.xlsx> <table2.xlsx> <out.xlsx>
```

The script joins the two tables horizontally and aligns rows based on their mileage values (column C of the first table and column B of the second table). Rows are merged when the mileage difference does not exceed `0.1`; otherwise, blank rows are inserted so that both sides remain aligned. The result is written to `<out.xlsx>`.
