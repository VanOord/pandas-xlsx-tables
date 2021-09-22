from typing import BinaryIO, Optional, Union

from openpyxl import Workbook
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet
from pandas import DataFrame


def df_to_table(
    df: DataFrame,
    table_name: str,
    ws: Worksheet,
    include_index: bool = True,
    include_header: bool = True,
    table_style: Optional[TableStyleInfo] = None,
) -> Table:

    # note that this returns index and colum names on separate rows
    rows = dataframe_to_rows(df, index=include_index, header=include_header)

    if include_index:
        header = next(rows)
        if include_header:
            # Don't include the empy row returned here
            index_header = next(rows)
            header = [*index_header, *header[len(index_header) :]]
        ws.append([str(h) for h in header])
    elif include_header:
        # discard the entry for the index name, as we have no header row
        next(rows)
    for row in rows:
        ws.append(row)

    tab = Table(
        displayName=table_name,
        ref=(
            f"A{1 if include_header else 2}:"
            f"{get_column_letter(len(df.columns) + 1 if include_index else 0)}"
            f"{len(df)+1}"
        ),
    )

    # Use the default excel table style with striped rows and banded columns
    if not table_style:
        table_style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=include_index,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
    tab.tableStyleInfo = table_style
    ws.add_table(tab)
    return ws


def df_to_xlsx_table(
    df: DataFrame,
    table_name: str,
    file: Optional[Union[str, BinaryIO]],
    include_index: bool = True,
    include_header: bool = True,
    table_style: Optional[TableStyleInfo] = None,
) -> None:
    if not file:
        file = table_name + ".xlsx"

    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title=table_name)
    df_to_table(
        df,
        table_name,
        ws,
        include_index=include_index,
        include_header=include_header,
        table_style=table_style,
    )

    wb.save(file)
    return
