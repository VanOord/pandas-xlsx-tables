from typing import Dict, Literal

import pandas as pd
from pandas.api.types import (
    is_datetime64_any_dtype,
    is_float_dtype,
    is_integer_dtype,
    is_string_dtype,
)
from pandas.io.formats.format import FloatArrayFormatter

NamedTableStyle = Literal[
    "Table Style Dark 1",
    "Table Style Dark 2",
    "Table Style Dark 3",
    "Table Style Dark 4",
    "Table Style Dark 5",
    "Table Style Dark 6",
    "Table Style Dark 7",
    "Table Style Dark 8",
    "Table Style Dark 9",
    "Table Style Dark 10",
    "Table Style Dark 11",
    "Table Style Light 1",
    "Table Style Light 2",
    "Table Style Light 3",
    "Table Style Light 4",
    "Table Style Light 5",
    "Table Style Light 6",
    "Table Style Light 7",
    "Table Style Light 8",
    "Table Style Light 9",
    "Table Style Light 10",
    "Table Style Light 11",
    "Table Style Light 12",
    "Table Style Light 13",
    "Table Style Light 14",
    "Table Style Light 15",
    "Table Style Light 16",
    "Table Style Light 17",
    "Table Style Light 18",
    "Table Style Light 19",
    "Table Style Light 20",
    "Table Style Light 21",
    "Table Style Medium 1",
    "Table Style Medium 2",
    "Table Style Medium 3",
    "Table Style Medium 4",
    "Table Style Medium 5",
    "Table Style Medium 6",
    "Table Style Medium 7",
    "Table Style Medium 8",
    "Table Style Medium 9",
    "Table Style Medium 10",
    "Table Style Medium 11",
    "Table Style Medium 12",
    "Table Style Medium 13",
    "Table Style Medium 14",
    "Table Style Medium 15",
    "Table Style Medium 16",
    "Table Style Medium 17",
    "Table Style Medium 18",
    "Table Style Medium 19",
    "Table Style Medium 20",
    "Table Style Medium 21",
    "Table Style Medium 22",
    "Table Style Medium 23",
    "Table Style Medium 24",
    "Table Style Medium 25",
    "Table Style Medium 26",
    "Table Style Medium 27",
    "Table Style Medium 28",
]


def create_format_mapping(workbook):
    return {
        "text": workbook.add_format({"num_format": "@"}),
        "float1": workbook.add_format({"num_format": "0.0"}),
        "float2": workbook.add_format({"num_format": "0.00"}),
        "float3": workbook.add_format({"num_format": "0.000"}),
        "float6": workbook.add_format({"num_format": "0.000000"}),
        "int": workbook.add_format({"num_format": "0"}),
        "datetime-milliseconds": workbook.add_format(
            {"num_format": "yyyy-mm-dd hh:mm:ss.000"}
        ),
        "datetime-seconds": workbook.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"}),
        "datetime-minutes": workbook.add_format({"num_format": "yyyy-mm-dd hh:mm"}),
        "date": workbook.add_format({"num_format": "yyyy-mm-dd"}),
        "scientific": workbook.add_format({"num_format": "0,00E+00"}),
    }


def format_for_col(col: pd.Series, format_mapping: Dict):
    # https://pandas.pydata.org/pandas-docs/stable/user_guide/timeseries.html#offset-aliases
    if is_integer_dtype(col):
        return format_mapping["int"]

    elif is_float_dtype(col):
        # use pandas internal function to determine float format (number of
        # digits or scientific notation)
        suffix = (
            FloatArrayFormatter(col[~col.isna()]).get_result_as_array()[0].split(".")[1]
        )
        if "e" in suffix:
            return format_mapping["scientific"]
        digits = len(suffix)
        if digits <= 1:
            return format_mapping["float1"]
        elif digits <= 2:
            return format_mapping["float2"]
        elif digits <= 3:
            return format_mapping["float3"]
        else:
            return format_mapping["float6"]

    elif is_string_dtype(col):
        return format_mapping["text"]
    elif is_datetime64_any_dtype(col):
        # check smallest used time unit for precision
        if all((col - col.dt.floor("1D")) == pd.Timedelta(0)):
            return format_mapping["date"]
        elif all((col - col.dt.floor("min")) == pd.Timedelta(0)):
            return format_mapping["datetime-minutes"]
        elif all((col - col.dt.floor("S")) == pd.Timedelta(0)):
            return format_mapping["datetime-seconds"]
        else:
            return format_mapping["datetime-milliseconds"]
