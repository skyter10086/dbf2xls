from openpyxl.worksheet.worksheet import Worksheet


def fill(ws: Worksheet, start_cell: str, data_list: list[list]):
    start_row = int(start_cell[1:])
    start_col = ord(start_cell[0]) - ord("A") + 1

    for i, row_data in enumerate(data_list):
        for j, cell_value in enumerate(row_data):
            ws.cell(row=start_row + i, column=start_col + j, value=cell_value)
