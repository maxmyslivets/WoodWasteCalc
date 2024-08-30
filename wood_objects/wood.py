import openpyxl

from errors.parse_errors import UnknownSpecie
from parsing.density_db import Density
from parsing.volumes_table import Volume


class Trunk:
    """
    Класс для создания объекта Ствол
    """
    volume: float

    def __init__(self, diameter: int, height: int) -> None:
        self.diameter = diameter
        self.height = height

    def __repr__(self) -> str:
        return f"Диаметр {self.diameter} см, высота {self.height} м"


class Wood:
    """
    Класс для создания объекта Дерево
    """

    def __init__(self, name: str, number: str, specie: str, is_shrub: bool, trunks: list[Trunk],
                 area: float = None) -> None:
        self.name = name
        self.number = number
        self.specie = specie
        self.is_shrub = is_shrub
        self.trunks = trunks
        self.area = area

    def __repr__(self) -> str:
        area = f", {self.area} м2" if self.area else ""
        if len(self.trunks) > 1:
            wood_str = f"{self.number}, {self.name}, {'кустарник' if self.is_shrub else 'дерево'}, " \
                       f"{len(self.trunks)} стволов"
            trunks_str = "\n".join([f"Ствол {idx + 1}: {trunk}" for idx, trunk in enumerate(self.trunks)])
            return f"{wood_str}\n{trunks_str}"
        else:
            return f"{self.number}, {self.name}, {'кустарник' if self.is_shrub else 'дерево'}, {self.trunks[0]}" + area


class WoodWaste(Wood):
    """
    Класс для расчёта отходов.
    """
    density: float
    data: list

    def __init__(self, name: str, number: str, specie: str, is_shrub: bool, trunks: list[Trunk],
                 area: float = None):
        super().__init__(name, number, specie, is_shrub, trunks, area)

        self.calculation()

    def calculation(self) -> None:
        # нахождение объёмов
        volume = Volume()
        for trunk in self.trunks:
            trunk.volume = volume.get_volume(trunk.diameter, trunk.height)
        # перемещение ствола с наибольшим объёмом в начало списка
        max_trunk = max(self.trunks, key=lambda t: t.volume)
        self.trunks.remove(max_trunk)
        self.trunks.insert(0, max_trunk)
        # определение плотности породы
        try:
            self.density = Density().get_density(self.specie)
        except UnknownSpecie as e:
            print(e)
            group = int(input(f"К какой группе плотности относится пород '{self.specie}'?\n"
                              f"1-мягкие, 2-твердые, 3-лиственница. Введите цифру: "))  # TODO: try-except
            group_name = {1: "мягкие", 2: "твердые", 3: "лиственница"}[group]
            Density().add_specie_to_group(group_name, self.specie)
            self.density = Density().get_density(self.specie)

    def export_preparation(self) -> None:
        self.data = []
        if not self.is_shrub:
            for idx, trunk in enumerate(self.trunks):
                self.data.append(
                    [self.number if idx == 0 else None, self.name if idx == 0 else None,
                     1, trunk.diameter, trunk.height, trunk.volume, self.density]
                )
        else:
            self.data.append(
                [self.number, self.name, 3 if self.area else 1, self.trunks[0].diameter, self.trunks[0].height,
                 self.trunks[0].volume, self.area if self.area else 1 / 3, self.density]
            )

    @staticmethod
    def export_to_xls(data, path) -> None:
        # Создаем новую рабочую книгу и выбираем активный лист
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # Переменная для отслеживания текущей строки в Excel
        current_row = 1

        # Проходим по внешнему списку
        for group in data:
            for row in group:
                # Записываем данные в строки листа Excel
                sheet.append(row)
                current_row += 1

        # Сохраняем рабочую книгу в файл
        workbook.save(path)
        print(f"Успешно {path}.")

    def __repr__(self) -> str:
        if not self.is_shrub:
            s = f"{self.number} | {self.name} | 1 | {self.trunks[0].diameter} | {self.trunks[0].height} | " \
                f"{self.trunks[0].volume} | {self.density}"
            s_ = ""
            for trunk in self.trunks[1:]:
                s_ += f"\n\t|\t| 1 | {trunk.diameter} | {trunk.height} | {trunk.volume} | {self.density}"
            return s + s_
        else:
            s = f"{self.number} | {self.name} | {3 if self.area else 1} | {self.trunks[0].diameter} | " \
                f"{self.trunks[0].height} | {self.trunks[0].volume} | {self.area if self.area else 1 / 3} | " \
                f"{self.density}"
            return s
