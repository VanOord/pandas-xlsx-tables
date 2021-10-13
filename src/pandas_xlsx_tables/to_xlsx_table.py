from typing import BinaryIO, Iterable, Literal, Optional, Tuple, Union

import numpy as np
import xlsxwriter
from openpyxl.worksheet.table import TableStyleInfo
from pandas import DataFrame

from .utils import NamedTableStyle, create_format_mapping, format_for_col

HeaderOrientation = Literal["diagonal", "horizontal", "vertical"]


def dfs_to_xlsx_tables(
    input: Iterable[Tuple[DataFrame, str]],
    file: Union[str, BinaryIO],
    index: bool = True,
    table_style: Optional[NamedTableStyle] = "Table Style Medium 9",
    nan_inf_to_errors=False,
    header_orientation: HeaderOrientation = "horizontal",
) -> None:
    """Convert multiple dataframes to an excel file.

    Args:
        input (Iterable[Tuple[DataFrame, str]]): A list of tuples of (df, table_name)
        file (Union[str, BinaryIO]): File name or descriptor for the output
        index (bool, optional): Include the datafrme index in the results. Defaults
            to True
        table_style (Optional[NamedTableStyle], optional): Excel table style. Defaults
            to "Table Style Medium 9".
        nan_inf_to_errors (bool, optional): Explicitly write nan/inf values as errors.
            Defaults to False.
        header_orientation (HeaderOrientation, optional): Rotate the table headers, can
            be horizontal, vertical or diagonal. Defaults to "horizontal".
    """
    wb = xlsxwriter.Workbook(file, options=dict(nan_inf_to_errors=nan_inf_to_errors))

    format_mapping = create_format_mapping(wb)
    if header_orientation == "diagonal":
        header_format = wb.add_format()
        header_format.set_rotation(45)
    elif header_orientation == "vertical":
        header_format = wb.add_format()
        header_format.set_rotation(90)

    for df, table_name in input:
        ws = wb.add_worksheet(name=table_name)
        if index:
            df = df.reset_index()
        if not nan_inf_to_errors:
            df = (
                df.replace(np.Inf, np.finfo(np.float64).max)
                .replace(-np.Inf, np.finfo(np.float64).min)
                .fillna("")
            )
        options = {
            "data": df.values,
            "name": table_name,
            "style": table_style,
            "first_column": index,
            "columns": [
                {"header": c, "format": format_for_col(df[c], format_mapping)}
                for c in df.columns
            ],
        }
        ws.add_table(0, 0, len(df), len(df.columns) - 1, options)
        if header_orientation == "diagonal":
            ws.set_row(
                0, max(15, 12 + 4 * max(len(c) for c in df.columns)), header_format
            )
        elif header_orientation == "vertical":
            ws.set_row(
                0, max(15, 4 + 6 * max(len(c) for c in df.columns)), header_format
            )
        elif header_orientation == "horizontal":
            # adjust row widths
            for i, width in enumerate([len(x) for x in df.columns]):
                ws.set_column(i, i, max(8.43, width))
    wb.close()
    return


def df_to_xlsx_table(
    df: DataFrame,
    table_name: str,
    file: Optional[Union[str, BinaryIO]] = None,
    index: bool = True,
    table_style: Optional[TableStyleInfo] = "Table Style Medium 9",
    nan_inf_to_errors=False,
    header_orientation: HeaderOrientation = "horizontal",
) -> None:
    """Convert single dataframe to an excel file.

    Args:
        df (DataFrame): Padas dataframe to convert to excel.
        table_name (str):Name of the table.
        file (Union[str, BinaryIO]): File name or descriptor for the output.
            Defaults to <table_name>.xlsx
        index (bool, optional): Include the datafrme index in the results. Defaults
            to True
        table_style (Optional[NamedTableStyle], optional): Excel table style. Defaults
            to "Table Style Medium 9".
        nan_inf_to_errors (bool, optional): Explicitly write nan/inf values as errors.
            Defaults to False.
        header_orientation (HeaderOrientation, optional): Rotate the table headers, can
            be horizontal, vertical or diagonal. Defaults to "horizontal".
    """
    dfs_to_xlsx_tables(
        [(df, table_name)],
        file=file or table_name + ".xlsx",
        index=index,
        table_style=table_style,
        nan_inf_to_errors=nan_inf_to_errors,
        header_orientation=header_orientation,
    )
