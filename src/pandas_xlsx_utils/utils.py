from typing import Tuple

from openpyxl.utils.cell import get_column_letter


def tuple_to_coordinate(tup: Tuple[int, int]):
    """Tuple of row,col to excel coordiante

    Inverse of openpyxl.utils.cell.coordinate_to_tuple
    tuple_to_coordinate(1,3) = "A3"

    """
    row, col = tup
    return f"{get_column_letter(col)}{row}"
