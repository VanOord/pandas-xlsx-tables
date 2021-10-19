# pandas-xlsx-tables

 Even though you might not like it, Excel isn't going anywhere. And Excel with tables is [a lot better than without](https://www.ecosia.org/search?q=advantages+of+excel+tables).

Strangely Pandas does not support reading from and writing to excel tables out of the box, and due to the complexity of Pandas this is [not easily added](https://github.com/pandas-dev/pandas/issues/24862) (though having this built into Pandas would be the prefered solution). This separate package is thus a separate companion to Pandas, with utility functions to read and write Excel Tables from and to Pandas DataFrames.

!["Excel screenshot](https://raw.githubusercontent.com/VanOord/pandas-xlsx-tables/master/docs/_static/xlsx_table.png)

```python
>>> from pandas_xlsx_tables import xlsx_table_to_df
>>> df = xlsx_table_to_df("my_file.xlsx", "Table1")
>>> df
     col1 col2
Row
0       1    a
1       2    b
```
And the reverse process:

```python
>>> from pandas_xlsx_tables import df_to_xlsx_table
>>> df_to_xlsx_table(df, "my_table", header_orientation="diagonal", index=False)
```

!["Excel screenshot](https://raw.githubusercontent.com/VanOord/pandas-xlsx-tables/master/docs/_static/xlsx_table_2.png)


## Contents

* [Overview](readme)
* [License](license)
* [Authors](authors)
* [Changelog](changelog)
* [Module Reference](api/modules)


## Indices and tables

```eval_rst
* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
```
