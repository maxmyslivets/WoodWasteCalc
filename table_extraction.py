from parsing.config import Config
from parsing.directory import get_files_in_directory
from parsing.dxf import dxf_parse
from xls.xls_write import xls_write


def table_extraction_from_dxf():

    config = Config()

    # экспорт таблиц из dxf в xls
    for dxf_file in get_files_in_directory(config.settings.dxf_directory):
        data, xls_filename = dxf_parse(dxf_file)
        data.insert(0, config.settings.input_table_structure)
        xls_write(data, config.settings.xls_directory / (xls_filename + '.xlsx'))


if __name__ == '__main__':

    table_extraction_from_dxf()
