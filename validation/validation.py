import json
from pathlib import Path

from errors.parse_errors import ParseQuantityError
from config.config import Config
from parsing.parse_xls import XLSWoodParser, RawWood
from taxation.species_db import Species


config = Config()


class ValidXLS:
    """
    Класс для проверки XLS файла на корректность
    """
    def __init__(self, filepath: Path) -> None:
        """
        :param filepath: путь к файлу XLS
        """
        self.file = filepath
        self.shrub_species, self.wood_species = Species().get_species()
        with (open(config.taxation_characteristics.density_json_path, 'r', encoding='utf-8') as file):
            data = json.load(file)
            group_of_wood = data.get('group_of_wood', {})
            self.density_species = []
            for _, species in group_of_wood.items():
                self.density_species.extend(species)

    def check_specie(self, specie: str) -> bool:
        """
        Проверка существования породы в файле species.json
        :param specie: название породы
        """
        return False if specie not in self.shrub_species and specie not in self.wood_species else True

    def check_density(self, specie: str) -> bool:
        """
        Проверка существования породы в файле densities.json
        :param specie: название породы
        """
        return False if specie not in self.density_species else True

    def check_valid(self) -> bool:
        """
        Проверка на корректность данный в файле XLS
        """

        print(f"Валидация файла `{self.file.name}` ...")

        is_valid = True

        table = XLSWoodParser().parse(self.file)

        species = []
        for row in table:
            species.append(row[1].split()[0].lower())

        # обнаружение неизвестных пород
        unknown_specie = []
        for specie in species:
            if not self.check_specie(specie) and specie not in unknown_specie:
                unknown_specie.append(specie)
        if len(unknown_specie) != 0:
            print(f"`{self.file.name}` - Обнаружены неизвестные породы: {unknown_specie}")
            is_valid = False

        # обнаружение пород с неизвестной плотностью древесины
        unknown_density = []
        for specie in species:
            if not self.check_density(specie) and specie not in unknown_density:
                unknown_density.append(specie)
        if len(unknown_density) != 0:
            print(f"`{self.file.name}` - Обнаружены породы с неизвестной плотностью древесины: {unknown_density}")
            is_valid = False

        for row in table:
            raw_wood = RawWood(*row)

            # парсинг номера
            try:
                raw_wood.parse_number()
            except Exception as e:
                print(f"`{self.file.name}` [{raw_wood}] - Ошибка извлечения списка номеров. {e}")
                is_valid = False

            # парсинг породы
            try:
                raw_wood.parse_specie()
            except Exception as e:
                print(f"`{self.file.name}` [{raw_wood}] - Ошибка извлечения породы. {e}")
                is_valid = False

            # парсинг количества
            try:
                raw_wood.parse_quantity()
                try:
                    if raw_wood.quantity_is_area and len(raw_wood.number) > 1:
                        raise ParseQuantityError(f"Количество номеров превышает 1 для площади.")
                except AttributeError:
                    # print(f"`{self.file.name}` [{raw_wood}] - Невозможно произвести проверку "
                    #       f"из-за ошибки в извлечении списка номеров")
                    is_valid = False
            except Exception as e:
                print(f"`{self.file.name}` [{raw_wood}] - Ошибка извлечения количества/площади. {e}")
                is_valid = False

            # парсинг диаметра
            try:
                raw_wood.parse_diameter()
            except AttributeError as e:
                # print(f"`{self.file.name}` [{raw_wood}] - Невозможно произвести проверку "
                #       f"из-за ошибки в извлечении предыдущих данных. {e}")
                is_valid = False
            except Exception as e:
                print(f"`{self.file.name}` [{raw_wood}] - Ошибка извлечения диаметра. {e}")
                is_valid = False

            # парсинг высоты
            try:
                raw_wood.parse_height()
            except AttributeError as e:
                # print(f"`{self.file.name}` [{raw_wood}] - Невозможно произвести проверку "
                #       f"из-за ошибки в извлечении предыдущих данных. {e}")
                is_valid = False
            except Exception as e:
                print(f"`{self.file.name}` [{raw_wood}] - Ошибка извлечения высоты. {e}")
                is_valid = False

            # сравнение количества диаметров и высот
            try:
                if len(raw_wood.height) != len(raw_wood.diameter) and '-' not in raw_wood._diameter:
                    print(f"`{self.file.name}` [{raw_wood}] - Неравное количество диаметров и высот.")
                    is_valid = False
            except Exception as e:
                is_valid = False

        return is_valid
