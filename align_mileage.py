import sys
import pandas as pd

TOLERANCE = 0.1


def read_excel(path):
    """Read workbook and return DataFrame."""
    return pd.read_excel(path)


def align_by_mileage(path1, path2, out_path):
    df1 = read_excel(path1)
    df2 = read_excel(path2)

    # column indexes: table1 use column C (index 2), table2 use column B (index 1)
    col1 = df1.columns[2] if len(df1.columns) > 2 else df1.columns[-1]
    col2 = df2.columns[1] if len(df2.columns) > 1 else df2.columns[0]

    i = j = 0
    rows = []
    cols = list(df1.columns) + list(df2.columns)

    while i < len(df1) or j < len(df2):
        row1 = df1.iloc[i] if i < len(df1) else None
        row2 = df2.iloc[j] if j < len(df2) else None

        m1 = row1[col1] if row1 is not None else None
        m2 = row2[col2] if row2 is not None else None

        try:
            m1_val = float(m1)
        except (TypeError, ValueError):
            m1_val = None
        try:
            m2_val = float(m2)
        except (TypeError, ValueError):
            m2_val = None

        if (m1_val is not None and m2_val is not None and
                abs(m1_val - m2_val) <= TOLERANCE):
            rows.append(list(row1) + list(row2))
            i += 1
            j += 1
        elif j >= len(df2) or (m1_val is not None and (m2_val is None or m1_val < m2_val)):
            rows.append(list(row1) + [pd.NA] * len(df2.columns))
            i += 1
        else:
            rows.append([pd.NA] * len(df1.columns) + list(row2))
            j += 1

    out_df = pd.DataFrame(rows, columns=cols)
    out_df.to_excel(out_path, index=False)


def main():
    if len(sys.argv) != 4:
        print("Usage: python align_mileage.py <table1.xlsx> <table2.xlsx> <out.xlsx>")
        sys.exit(1)
    align_by_mileage(sys.argv[1], sys.argv[2], sys.argv[3])


if __name__ == "__main__":
    main()
