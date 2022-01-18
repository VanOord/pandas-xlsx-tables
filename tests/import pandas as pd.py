import pandas as pd

from pandas_xlsx_tables import df_to_xlsx_table

df = pd.DataFrame([[pd.Timestamp(2022, 1, 22, 14, 15, 0), 1, 1.0, "b"]])
print(df)

df_to_xlsx_table(df, "my_table", header_orientation="diagonal", index=False)
