from parsing.config import Config
from parsing.directory import get_files_in_directory
from parsing.parse_xls import XLSParser, RawWood
from wood_objects.wood import WoodWaste


config = Config()


def main() -> None:

    xls_files = get_files_in_directory(config.directories.xls_directory)
    for file in xls_files:
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

        if not validate:
            print(f"Файл `{file.name}` имеет нераспознанные данные и не будет обработан")
            continue

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

        WoodWaste.export_to_xls(result_wood,
                                config.directories.out_directory / f"{file.stem}_out_wood{file.suffix}")
        WoodWaste.export_to_xls(result_shrub,
                                config.directories.out_directory / f"{file.stem}_out_shrub{file.suffix}")
