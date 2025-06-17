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
