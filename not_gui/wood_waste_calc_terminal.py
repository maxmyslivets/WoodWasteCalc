from config.config import Config
from validation.validation import ValidXLS
from parsing.parse_xls import XLSWoodParser, RawWood
from calculation.waste import WasteCalculation


config = Config()


def main() -> None:

    xls_files = [file for file in config.directories.xls_directory.iterdir() if file.is_file()]
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
            raw_woods = XLSWoodParser().parse(file)
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
            w = WasteCalculation(wood.name, wood.number, wood.specie, wood.is_shrub, wood.trunks, wood.area)
            w.export_preparation()
            if w.is_shrub:
                result_shrub.append(w.data)
            else:
                result_wood.append(w.data)

        if len(result_wood) > 1:
            WasteCalculation.export_to_xls(result_wood,
                                           config.directories.out_directory / f"{file.stem}_out_wood{file.suffix}")
        if len(result_shrub) > 1:
            WasteCalculation.export_to_xls(result_shrub,
                                           config.directories.out_directory / f"{file.stem}_out_shrub{file.suffix}")
