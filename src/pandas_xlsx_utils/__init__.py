import sys

if sys.version_info[:2] >= (3, 8):
    # TODO: Import directly (no need for conditional) when `python_requires = >= 3.8`
    from importlib.metadata import PackageNotFoundError, version  # pragma: no cover
else:
    from importlib_metadata import PackageNotFoundError, version  # pragma: no cover

try:
    # Change here if project is renamed and does not equal the package name
    dist_name = "pandas-xlsx-utils"
    __version__ = version(dist_name)
except PackageNotFoundError:  # pragma: no cover
    __version__ = "unknown"
finally:
    del version, PackageNotFoundError

from .from_xlsx_tables import xlsx_tables_to_frames
from .to_xlsx_table import frame_to_xlsx_table, frames_to_xlsx_tables

__all__ = ["frame_to_xlsx_table", "frames_to_xlsx_tables", "xlsx_tables_to_frames"]
