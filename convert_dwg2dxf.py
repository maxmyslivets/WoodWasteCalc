from parsing.config import Config
from dwg.converter import check_converter_in_system, install_converter, dwg2dxf


def convert_dwg_to_dxf() -> None:

    config = Config()

    # проверка наличия в системе конвертера dwg в dxf
    if not check_converter_in_system():
        install_converter()

    # конвертация dwg в dxf
    dwg2dxf(config.directories.dwg_directory, config.directories.dxf_directory)


if __name__ == '__main__':

    convert_dwg_to_dxf()
