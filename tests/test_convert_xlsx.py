import os
from datetime import datetime
from tempfile import TemporaryDirectory

import numpy as np
import pandas as pd
import pytest

from pandas_xlsx_tables import df_to_xlsx_table, dfs_to_xlsx_tables, xlsx_tables_to_dfs


@pytest.fixture
def cleandir():
    with TemporaryDirectory() as newpath:
        old_cwd = os.getcwd()
        os.chdir(newpath)
        yield
        os.chdir(old_cwd)


@pytest.fixture()
def df():
    t1 = datetime(2021, 1, 2, 12, 13, 14)
    t2 = datetime(2020, 2, 1, 10, 56, 16)
    df = pd.DataFrame(
        [
            ["Apples", 10000, "a", "nan", t1, datetime(2021, 1, 2)],
            ["Pears", 2000, "p", 12e-12, t1, datetime(2021, 2, 2)],
            ["Bananas", 6000, "", 6500, t2, datetime(2021, 1, 3)],
            ["Oranges", 500, "o", "inf", t2, datetime(2021, 4, 2)],
            ["Plums", 500, "o", -np.Inf, t2, datetime(2021, 4, 2)],
        ],
        columns=["name", "int", "A very long header", "scientific", "datetime", "date"],
    ).set_index("name")
    df.scientific = df.scientific.astype(float)
    df.date = pd.to_datetime(df.date)
    return df


@pytest.mark.usefixtures("cleandir")
class TestToXlsx:
    def test_write_df_to_xlsx_table(self, df):
        df_to_xlsx_table(df, "TestTable", "test.xlsx")

    def test_write_multiple_dfs_to_xlsx_table(self, df):
        dfs_to_xlsx_tables(
            (
                (df, "TestTable1"),
                (df, "TestTable2"),
                (df, "TestTable3"),
            ),
            "test_multiple.xlsx",
        )


@pytest.mark.usefixtures("cleandir")
class TestRoundtrip:
    @pytest.mark.parametrize("nan_inf_to_errors", (True, False))
    @pytest.mark.parametrize("index", (True, False))
    def test_roundtrip_single_table(self, df, nan_inf_to_errors, index):

        if not index:
            df = df.reset_index(drop=True)
        df_to_xlsx_table(
            df,
            "single_table",
            nan_inf_to_errors=nan_inf_to_errors,
            index=index,
        )

        result = xlsx_tables_to_dfs("single_table.xlsx")["single_table"]
        assert df.columns.equals(result.columns)
        assert df.index.equals(result.index)

        # nan and inf do not compare as equal, so use this work around:
        r = {
            np.Inf: "_inf_",
            -np.Inf: "_-inf_",
            np.NaN: "_nan_",
        }

        if nan_inf_to_errors:
            # Xlwxwriter does not discern +/- infinity with this option
            r[-np.Inf] = r[np.Inf]
        assert df.replace(r).equals(result.replace(r))

    @pytest.mark.parametrize("nan_inf_to_errors", (True, False))
    @pytest.mark.parametrize("index", (True, False))
    @pytest.mark.parametrize(
        "columns",
        (
            ["int"],
            ["int", "A very long header", "scientific", "datetime", "date"],
        ),
    )
    def test_roundtrip_multiple_tables(self, df, nan_inf_to_errors, index, columns):
        df = df[columns]
        if not index:
            df = df.reset_index(drop=True)
        input = ((df, "TestTable1"), (df, "TestTable2"), (df, "TestTable3"))
        dfs_to_xlsx_tables(
            input,
            "test_multiple.xlsx",
            nan_inf_to_errors=nan_inf_to_errors,
            index=index,
        )

        result = xlsx_tables_to_dfs("test_multiple.xlsx")["TestTable2"]
        assert df.columns.equals(result.columns)
        assert df.index.equals(result.index)

        # nan and inf do not compare as equal, so use this work around:
        r = {
            np.Inf: "_inf_",
            -np.Inf: "_-inf_",
            np.NaN: "_nan_",
        }

        if nan_inf_to_errors:
            # Xlwxwriter does not discern +/- infinity with this option
            r[-np.Inf] = r[np.Inf]
        assert df.replace(r).equals(result.replace(r))

    def test_write_dates_with_timezone(self):
        df = pd.DataFrame(
            pd.date_range("2021-10-19T21:30:00Z", "2021-10-19T23:30:00Z", freq="30min"),
            columns=["dates"],
        )
        with pytest.raises(
            TypeError, match="Excel doesn't support timezones in datetimes"
        ):
            df_to_xlsx_table(df, "test_write_dates_with_timezone")
        df_to_xlsx_table(df, "test_write_dates_with_timezone", remove_timezone=True)

        result = xlsx_tables_to_dfs("test_write_dates_with_timezone.xlsx")[
            "test_write_dates_with_timezone"
        ]
        result["dates"] = result["dates"].dt.tz_localize("UTC")
        assert df.equals(result)

    def test_write_numeric_header(self):
        df = pd.DataFrame([[pd.Timestamp(2022, 1, 22, 14, 15, 0), 1, 1.0, "b"]])
        df_to_xlsx_table(df, "test_write_numeric_header")
        result = xlsx_tables_to_dfs("test_write_numeric_header.xlsx")[
            "test_write_numeric_header"
        ]
        assert (df.values == result.values).all()
