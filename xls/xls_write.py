from openpyxl import Workbook


def xls_write(data: list, path: str) -> None:

    wb = Workbook()
    ws = wb.active
    ws.title = "Table Data"

    # Записываем данные
    for i, row in enumerate(data, start=1):
        for j, value in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=value)

    wb.save(path)
