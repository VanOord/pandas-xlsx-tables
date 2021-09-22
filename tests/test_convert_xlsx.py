import os
from tempfile import TemporaryDirectory

import pandas as pd
import pytest

from pandas_xlsx_utils import (
    frame_to_xlsx_table,
    frames_to_xlsx_tables,
    xlsx_tables_to_frames,
)


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
class TestToXlsx:
    def test_write_df_to_xlsx_table(self, df):
        frame_to_xlsx_table(df, "TestTable", "test.xlsx")

    def test_write_multiple_dfs_to_xlsx_table(self, df):
        frames_to_xlsx_tables(
            (
                (df, "TestTable1"),
                (df, "TestTable2"),
                (df, "TestTable3"),
            ),
            "test_multiple.xlsx",
        )


@pytest.mark.usefixtures("cleandir")
class TestFromXlsx:
    def test_read_multiple_tables_from_xlsx(self, df):
        input = ((df, "TestTable1"), (df, "TestTable2"), (df, "TestTable3"))
        frames_to_xlsx_tables(
            input,
            "test_multiple.xlsx",
        )
        assert df.equals(xlsx_tables_to_frames("test_multiple.xlsx")["TestTable1"])
        assert df.equals(xlsx_tables_to_frames("test_multiple.xlsx")["TestTable2"])
        assert df.equals(xlsx_tables_to_frames("test_multiple.xlsx")["TestTable3"])
