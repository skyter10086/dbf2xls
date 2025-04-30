from typing import TypedDict
from openpyxl.worksheet.page import PageMargins

class PageSetup(TypedDict):
    print_area: str
    title_rows: str
    orientation: str
    paper_size: int
    page_margins: PageMargins
    fit_page: bool
    fit2width: bool
    fit2height: bool
    header: str
    footer: str
    horizontal_centered: bool
    vertical_centered: bool

