from openpyxl.styles import Font, Alignment, Border, Side
from typing import TypedDict

class Style(TypedDict):
    font: Font
    align: Alignment
    border: Border
    rows_height: list[int]
    cols_width: list[int]
    num_fmt: list[str] # 数组长度等于列数

