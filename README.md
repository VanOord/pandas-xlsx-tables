# pandas-xlsx-tables

Even though you might not like it, Excel isn't going anywhere. And Excel with tables is [a lot better than without](https://www.ecosia.org/search?q=advantages+of+excel+tables). Some highlights are: better performance, column references by name (instead of named ranges), sticky headers (instead of freeze panes), stricter typing (instead of random types), and sort/filter dropdowns.

By default, Pandas provides read and write functionality for Excel, but it cannot create native Excel tables (the result is only formatted like one). This is where pandas-xlsx-tables comes in: it converts Excel tables to dataframes and vice versa while mostly preserving data types. The API has been kept deliberately simple to provide useful functionality out of the box.

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


## Why not integrate this in Pandas directly?

Due to the complexity of Pandas and the large number of users, it is very difficult to significantly change the current excel implementation. Additionally, the available extension points for different engines are limited. Basically I tried and [gave up](https://github.com/pandas-dev/pandas/issues/24862) (but of course I would prefer having this built into Pandas). 
