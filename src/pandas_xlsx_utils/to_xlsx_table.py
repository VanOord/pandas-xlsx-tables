import warnings
from typing import BinaryIO, Iterable, Optional, Tuple, Union

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo
from pandas import DataFrame

from .utils import tuple_to_coordinate


def df_to_table(
    df: DataFrame,
    table_name: str,
    ws: WriteOnlyWorksheet,
    index: bool = True,
    table_style: Optional[TableStyleInfo] = None,
) -> Table:

    # note that this returns index and colum names on separate rows
    rows = dataframe_to_rows(df, index=index, header=True)

    header = next(rows)
    if index:
        # Don't include the empty row returned here
        index_header = next(rows)
        header = [*index_header, *header[len(index_header) :]]

    ws.append([str(h) for h in header])
    for row in rows:
        ws.append(row)

    bottom_right = (
        len(df.columns) + 1 if index else 0,
        len(df) + 1,
    )
    tab = Table(
        displayName=table_name,
        ref=f"A1:{tuple_to_coordinate(bottom_right)}",
        tableColumns=[TableColumn(i + 1, name=str(h)) for i, h in enumerate(header)],
    )

    # Use the default excel table style with striped rows and banded columns
    if not table_style:
        table_style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=index,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
    tab.tableStyleInfo = table_style

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=UserWarning)
        ws.add_table(tab)
    return ws


def dfs_to_xlsx_tables(
    input: Iterable[Tuple[DataFrame, str]],
    file: Union[str, BinaryIO],
    index: bool = True,
    table_style: Optional[TableStyleInfo] = None,
) -> None:

    if not table_style:
        table_style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=index,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )

    wb = Workbook(write_only=True)
    for df, table_name in input:
        ws = wb.create_sheet(title=table_name)
        df_to_table(
            df,
            table_name,
            ws,
            index=index,
            table_style=table_style,
        )

    wb.save(file)
    return


def df_to_xlsx_table(
    df: DataFrame,
    table_name: str,
    file: Optional[Union[str, BinaryIO]],
    index: bool = True,
    table_style: Optional[TableStyleInfo] = None,
) -> None:

    dfs_to_xlsx_tables(
        [(df, table_name)],
        file=file or table_name + ".xlsx",
        index=index,
        table_style=table_style,
    )
