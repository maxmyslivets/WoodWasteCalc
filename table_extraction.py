from parsing.config import Config
from parsing.directory import get_files_in_directory
from parsing.dxf import dxf_parse
from xls.xls_write import xls_write


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
    for dxf_file in get_files_in_directory(config.directories.dxf_directory):
        data, xls_filename = dxf_parse(dxf_file)
        data.insert(0, columns)
        xls_write(data, config.directories.xls_directory / (xls_filename + '.xlsx'))


if __name__ == '__main__':

    table_extraction_from_dxf()
