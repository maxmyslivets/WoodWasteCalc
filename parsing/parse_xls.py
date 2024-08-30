import string
import re
from openpyxl import load_workbook, Workbook

from wood_objects.wood import Wood, Trunk
from errors.parse_errors import *

from .config import Config

config = Config()

from .species_db import read_species_from_json

shrub_species, wood_species = read_species_from_json(config.settings.species_json_path)


class RawWood:
    """
    Класс для создания списка деревьев из сырой строки.
    """

    def __init__(self, number: str, name: str, quantity: str, diameter: str, height: str) -> None:
        self._number = number
        self._name = name
        self._quantity = quantity
        self._diameter = diameter
        self._height = height

        # дополнительные переменные
        self.specie: str  # порода
        self.is_shrub: bool  # кустарник
        self.trunk_count: int  # количество стволов

    def __repr__(self) -> str:
        return f"'{self._number}'\t'{self._name}'\t'{self._quantity}'\t" \
               f"'{self._diameter}'\t'{self._height}'"

    def _parse_number(self) -> list[str]:
        """
        Форматирование номера дерева или деревьев в таблице.
        :return: Список номеров деревьев
        """
        result = []
        parts = self._number.translate({ord(c): None for c in string.whitespace}).split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                start, end = part.split('-')
                result.extend([str(_) for _ in range(int(start), int(end) + 1)])
            else:
                result.append(part)
        return result

    def _parse_specie(self) -> dict:
        """
        Парсинг строки наименования породы.
        :return: словарь значений {порода: str, кустарник: bool, количество стволов: int}
        """

        # Название породы
        specie = self._name.split()[0].lower()

        # Проверка наличия породы в списке известных
        if specie not in shrub_species and specie not in wood_species:
            raise UnknownSpecie("Неизвестная порода")

        # Определение: кустарник или дерево
        if specie in shrub_species or "(поросль)" in self._name:
            is_shrub = True
        elif specie in wood_species and "(поросль)" not in self._name:
            is_shrub = False
        else:
            raise ParseSpecieError("Не удалось определить кустарник или дерево")

        # Определение количества стволов
        match = re.search(r'(\d+)\s*ствол', self._name, re.IGNORECASE)
        # trunk_count = int(match.group(1)) if match else len(self._parse_number())
        trunk_count = int(match.group(1)) if match else 1

        return {"specie": specie, "is_shrub": is_shrub, "trunk_count": trunk_count}

    def _parse_quantity(self) -> (int, bool):
        """
        Определение количества деревьев или площади поросли.
        :return: (число: int, площадь: bool)
        """

        is_area = "м" in self._quantity

        quantity_match = re.search(r'\d+', self._quantity)
        # print(quantity_match)
        try:
            quantity = int(quantity_match.group(0))
        except Exception as e:
            raise ParseQuantityError(f"Ошибка в извлечении количества деревьев или площади.\n{e}")

        if is_area and len(self._parse_number()) > 1:
            raise ParseQuantityError(f"Ошибка в извлечении количества деревьев или площади.")

        return quantity, is_area

    def _parse_diameter(self) -> list[float]:
        """
        Форматирование списка диаметров
        :return: список диаметров
        """

        result = []

        try:
            # Замена кириллического "х" на латинское "x"
            diameters = self._diameter.replace("х", "x").replace("Х", "x")

            # Разделение строки по запятым
            parts = diameters.split(',')
            for part in parts:
                part = part.strip()
                # Поиск шаблона "числоxчисло"
                match = re.match(r'(\d+(\.\d+)?)x(\d+)', part)

                if match:
                    # Если найдено, распаковываем значения
                    num, _, repeat = match.groups()
                    result.extend([float(num)] * int(repeat))
                elif part.isdigit() or re.match(r'^\d+(\.\d+)?$', part):
                    # Если часть является числом, добавляем его в результат
                    result.append(float(part))
                elif part == "-":
                    # Специальный случай для "-"
                    result.append(2.0)

            number_count = len(self._parse_number())
            specie = self._parse_specie()
            is_area = self._parse_quantity()[1]
            if len(result) < number_count:
                if len(result) == 1 and not specie["is_shrub"]:
                    result *= number_count
                else:
                    raise ParseDiameterError(f"Ошибка парсинга строки диаметров стволов. Диаметров меньше чем номеров")
            if len(result) > number_count and not is_area:
                if not specie["trunk_count"] > number_count == 1:
                    raise ParseDiameterError(f"Ошибка парсинга строки диаметров стволов. Диаметров больше чем номеров")

        except Exception as e:
            raise ParseDiameterError(f"Ошибка парсинга строки диаметров стволов.\n{e}")

        return result

    def _parse_height(self) -> list[float]:
        """
        Форматирование списка высот
        :return: список высот
        """

        if self._parse_quantity()[0] == 1 and self._height.count(",") == 1:
            return [float(self._height.replace(',', '.'))]

        result = []

        try:
            # Замена кириллического "х" на латинское "x"
            heights = self._height.replace("х", "x").replace("Х", "x")

            # Разделение строки по запятым
            parts = heights.split(',')

            for part in parts:
                part = part.strip().lower()
                # Поиск шаблона "числоxчисло"
                match = re.match(r'(\d+(\.\d+)?)x(\d+)', part)

                if match:
                    # Если найдено, распаковываем значения
                    num, _, repeat = match.groups()
                    result.extend([float(num)] * int(repeat))
                elif part.isdigit() or re.match(r'^\d+(\.\d+)?$', part):
                    # Если часть является числом, добавляем его в результат
                    result.append(float(part))
                elif "свыше" in part:
                    # Обработка случая "свыше Xм"
                    number_match = re.search(r'\d+', part)
                    if number_match:
                        result.append(float(number_match.group()) + 0.5)
                elif "ниже" in part:
                    # Обработка случая "ниже Xм"
                    number_match = re.search(r'\d+', part)
                    if number_match:
                        h = float(number_match.group())
                        result.append(h - 0.5 if h > 0.5 else h)

            number_count = len(self._parse_number())
            specie = self._parse_specie()
            is_area = self._parse_quantity()[1]
            if len(result) < number_count:
                if len(result) == 1 and not specie["is_shrub"]:
                    result *= number_count
                else:
                    raise ParseDiameterError(f"Ошибка парсинга строки высот стволов. Высот меньше чем стволов")
            if len(result) > number_count and not is_area:
                if not specie["trunk_count"] > number_count == 1:
                    raise ParseDiameterError(f"Ошибка парсинга строки высот стволов. Высот больше чем номеров")

        except Exception as e:
            raise ParseHeightError(f"Ошибка парсинга строки высот стволов.\n{e}")

        return result

    def is_valid(self) -> None:
        """
        Валидация данных в строке
        :return:
        """
        valid = True
        try:
            # проверка парсинга
            number = self._parse_number()
            specie = self._parse_specie()
            quantity = self._parse_quantity()
            diameter = self._parse_diameter()
            height = self._parse_height()

            valid_list = {
                1: ("Ошибка проверки соответствия кустарнику по породе и количеству",
                    not (not specie["is_shrub"] and quantity[1])),
                # 2: ("Ошибка проверки равенства количества номеров и деревьев",
                #     len(number) == quantity[0] and not quantity[1]),
                3: ("Ошибка проверки количества номеров для кустарника с одной площадью",
                    not (len(number) > 1 and quantity[1]))
            }

            errors = []

            for error, valid_ in valid_list.values():
                if not valid_:
                    valid = False
                    errors.append(error)

            if errors:
                error = '\n'.join([_ for _ in errors])
                raise ValidError(f"Ошибка валидации в строке: {self.__repr__()}\n{error}")

        except Exception as e:
            raise ValidError(f"Ошибка валидации в строке: {self.__repr__()}\n{e}")

        return valid    # FIXME: Expected type 'None', got 'bool' instead

    def parse(self) -> list[Wood]:
        """

        :return:
        """

        woods = []

        number = self._parse_number()
        specie_dict = self._parse_specie()
        specie, is_shrub, trunk_count = specie_dict["specie"], specie_dict["is_shrub"], specie_dict["trunk_count"]
        quantity, is_area = self._parse_quantity()
        diameter = self._parse_diameter()
        height = self._parse_height()
        area = quantity if is_area else None
        try:
            for idx_wood in range(len(number)):
                trunks = []
                if trunk_count > 1:
                    for idx_trunk in range(trunk_count):
                        # FIXME: Expected type 'int', got 'float' instead
                        trunks.append(Trunk(diameter=diameter[idx_trunk], height=height[idx_trunk]))
                    woods.append(Wood(name=self._name, number=number[idx_wood], specie=specie, is_shrub=is_shrub,
                                      trunks=trunks))
                else:
                    # FIXME: Expected type 'int', got 'float' instead
                    trunks.append(Trunk(diameter=diameter[idx_wood], height=height[idx_wood]))
                    woods.append(Wood(name=self._name, number=number[idx_wood], specie=specie, is_shrub=is_shrub,
                                      trunks=trunks, area=area))

        except Exception as e:
            raise ParseError(f"Ошибка получения объекта дерева. {e}")

        return woods


class XLSParser:
    """
    Класс для чтения excel файлов.

    Входной файл должен иметь следующую структуру:
    1 строка — заголовки таблиц или пусто;
    A — номер;
    B — наименование;
    C — количество;
    D — диаметр;
    E — высота;
    """
    def __init__(self) -> None:

        self.columns = ['A', 'B', 'C', 'D', 'E']
        # TODO: добавить настройку структуры excel файла

    def parse(self, filepath: str) -> list:     # TODO: Try-Except
        """
        Парсинг excel файла. Получение двумерного массива данных таблицы.
        :param filepath: путь до файла
        :return: двумерный массив (list)
        """
        # Чтение исходного файла Excel
        wb = load_workbook(filepath)
        ws = wb.active

        raw_woods = []
        for row in range(2, ws.max_row + 1):
            wood_list = []
            for i, col in enumerate(self.columns, start=1):
                wood_list.append(str(ws[f'{col}{row}'].value))
            raw_woods.append(wood_list)

        return raw_woods

