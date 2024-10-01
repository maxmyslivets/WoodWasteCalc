import ezdxf
from openpyxl import Workbook

from config.config import Config


def dxf_parse(dxf_filename: str) -> tuple:

    doc = ezdxf.readfile(dxf_filename)

    mt_list = [[mt.plain_text().replace('\n', ' '), mt.dxf.insert[1], mt.dxf.insert[0]] for mt in
               doc.modelspace().query('MTEXT')]
    mt_list_sorted_by_x = sorted(mt_list, key=lambda m: m[1])[::-1]
    name_group = mt_list_sorted_by_x.pop(0)[0]

    n_coloumns = 5
    data_array = [mt_list_sorted_by_x[i:i + n_coloumns] for i in range(0, len(mt_list_sorted_by_x), n_coloumns)]
    for idx, row in enumerate(data_array):
        data_array[idx] = sorted(data_array[idx], key=lambda m: m[2])
    table_array = []
    for row in data_array:
        table_array.append([mt[0] for mt in row])

    return table_array, name_group


def xls_write(data: list, path: str) -> None:

    wb = Workbook()
    ws = wb.active
    ws.title = "Table Data"

    # Записываем данные
    for i, row in enumerate(data, start=1):
        for j, value in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=value)

    wb.save(path)


def table_extraction_from_dxf():

    config = Config()

    structure = (
        (config.table_structure.number, 'номер'),
        (config.table_structure.specie, 'порода'),
        (config.table_structure.quality, 'количество'),
        (config.table_structure.height, 'высота'),
        (config.table_structure.diameter, 'диаметр'),
    )

    sorted_structure = sorted(structure, key=lambda x: x[0])
    columns = [item[1] for item in sorted_structure]

    # экспорт таблиц из dxf в xls
    for dxf_file in [file for file in config.directories.dxf_directory.iterdir() if file.is_file()]:
        data, xls_filename = dxf_parse(dxf_file)
        data.insert(0, columns)
        xls_write(data, config.directories.xls_directory / (xls_filename + '.xlsx'))


if __name__ == '__main__':

    table_extraction_from_dxf()
