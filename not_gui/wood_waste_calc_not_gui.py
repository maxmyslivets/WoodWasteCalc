from parsing.config import Config
from parsing.directory import get_files_in_directory
from validation.validation import ValidXLS
from parsing.parse_xls import XLSParser, RawWood
from wood_objects.wood import WoodWaste


config = Config()


def main() -> None:

    xls_files = get_files_in_directory(config.directories.xls_directory)
    xls_files_valid = xls_files.copy()

    # валидация таблиц
    for file in xls_files:
        valid_xls = ValidXLS(file)
        is_valid = valid_xls.check_valid()
        if not is_valid:
            print(f"Файл `{file.name}` имеет нераспознанные данные и не будет обработан")
            xls_files_valid.remove(file)

    # обработка таблиц
    for file in xls_files_valid:
        try:
            raw_woods = XLSParser().parse(file)
        except Exception as e:
            print(f"`{file.name}` - Ошибка чтения данных из таблицы. {e}")
            continue
        woods = []
        try:
            for wood_list in raw_woods:
                raw_wood = RawWood(*wood_list)
                for wood in raw_wood.parse():
                    woods.append(wood)
        except Exception as e:
            print(f"`{file.name}` - Ошибка извлечения данных из таблицы. {e}")
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

        if len(result_wood) > 1:
            WoodWaste.export_to_xls(result_wood,
                                    config.directories.out_directory / f"{file.stem}_out_wood{file.suffix}")
        if len(result_shrub) > 1:
            WoodWaste.export_to_xls(result_shrub,
                                    config.directories.out_directory / f"{file.stem}_out_shrub{file.suffix}")
