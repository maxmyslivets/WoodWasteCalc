import sys
from pathlib import Path

from parsing.config import Config
from parsing.directory import list_files_in_directory
from parsing.dxf import dxf_parse
from xls.xls_write import xls_write
from parsing.parse_xls import XLSParser, RawWood
from wood_objects.wood import WoodWaste

config = Config()


def main() -> None:

    # Парсинг dxf файлов
    dxf_for_parse = list_files_in_directory(config.settings.parse_dxf_directory)
    xls_directory = Path(config.settings.parse_xls_directory)
    for dxf in dxf_for_parse:
        data, xls_filename = dxf_parse(dxf)
        xls_write(data, xls_directory/(xls_filename+'.xlsx'))

    xls_for_parse = list_files_in_directory(config.settings.parse_xls_directory)
    for file in xls_for_parse:
        raw_woods = XLSParser().parse(file)

        woods = []
        validate = True
        # Парсинг таблиц
        for wood_list in raw_woods:
            raw_wood = RawWood(*wood_list)
            try:
                if raw_wood.is_valid():
                    for wood in raw_wood.parse():
                        woods.append(wood)
            except Exception as e:
                validate = False
                print(e)

        # TODO: Обработка непрочитанных деревьев
        if not validate:
            if input("Имеются ошибки в некоторых исходных данных. Продолжать (пустой ввода, если да)?") != "":
                sys.exit(0)

        result_wood = [[["номер", "порода", "количество", "диаметр", "высота", "объем", "плотность", "стволы", "сучья",
                         "пни"]],]
        result_shrub = [[["номер", "порода", "количество", "диаметр", "высота", "объем", "площадь", "плотность",
                          "сучья", "пни"]],]
        for wood in woods:
            w = WoodWaste(wood.name, wood.number, wood.specie, wood.is_shrub, wood.trunks, wood.area)
            w.export_preparation()
            if w.is_shrub:
                result_shrub.append(w.data)
            else:
                result_wood.append(w.data)

        file_path = Path(file)
        file_out_wood = file_path.parent.parent / "out" / f"{file_path.stem}_out_wood{file_path.suffix}"
        file_out_shrub = file_path.parent.parent / "out" / f"{file_path.stem}_out_shrub{file_path.suffix}"

        WoodWaste.export_to_xls(result_wood, file_out_wood)
        WoodWaste.export_to_xls(result_shrub, file_out_shrub)
