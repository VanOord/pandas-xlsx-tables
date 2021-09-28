# pandas-xlsx-utils

If one accepts that Excel isn't going anywhere, the question becomes how to make the best of it. The best, but often underutilized feature of Excel is it's table feature: No more freeze panes, custom filters or issues when you sort one column but not the other.

Out of the box Pandas does not support reading and writing excel tables, and as the API of pandas is already pretty complex. So instead of adding a feature inside Pandas this separte package provides the required utility functions to read and write between Excel Tables and Pandas DataFrames.

!["Excel screenshot](docs/_static/xlsx_table.png)

```python
>>> from pandas_xlsx_utils import xlsx_table_to_frame
>>> xlsx_table_to_frame("my_file.xlsx", "Table1")
     col1 col2
Row
0       1    a
1       2    b

```
