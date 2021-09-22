import os
from tempfile import TemporaryDirectory

import pandas as pd
import pytest

from pandas_xlsx_utils import df_to_xlsx_table, dfs_to_xlsx_tables


@pytest.fixture
def cleandir():
    with TemporaryDirectory() as newpath:
        old_cwd = os.getcwd()
        os.chdir(newpath)
        yield
        os.chdir(old_cwd)


@pytest.fixture()
def df():
    return pd.DataFrame(
        [
            {"col1": 1, "col2": "a"},
            {"col1": 2, "col2": "b"},
        ],
    )


@pytest.mark.usefixtures("cleandir")
class TestDfToXlsx:
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
