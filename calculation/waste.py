import openpyxl

from wood.wood import Wood, Trunk
from taxation.density_db import Density
from taxation.volumes_table import Volume


class WasteCalculation(Wood):
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
            trunk.volume = volume.get_volume(int(trunk.diameter), int(trunk.height))
        # перемещение ствола с наибольшим объёмом в начало списка
        max_trunk = max(self.trunks, key=lambda t: t.volume)
        self.trunks.remove(max_trunk)
        self.trunks.insert(0, max_trunk)
        # определение плотности породы
        self.density = Density().get_density(self.specie)

    def export_preparation(self) -> None:
        self.data = []
        if not self.is_shrub:
            for idx, trunk in enumerate(self.trunks):
                self.data.append(
                    [self.number if idx == 0 else None, self.name if idx == 0 else None,
                     1, trunk.diameter, trunk.height, trunk.volume, self.density,
                     f"=1*{self.density}/1000*{trunk.volume}", f"=1*{self.density}/1000*{trunk.volume}*0.2",
                     f"=1*{self.density}/1000*{trunk.volume}*0.11" if idx == 0 else None]
                )
        else:
            self.data.append(
                [self.number, self.name, 3 if self.area else 1, self.trunks[0].diameter, self.trunks[0].height,
                 self.trunks[0].volume, self.area if self.area else '=1/3', self.density,
                 f"={3 if self.area else 1}*{self.trunks[0].volume}*{self.area if self.area else '1/3'}*"
                 f"{self.density/1000}", f"={3 if self.area else 1}*{self.trunks[0].volume}*"
                                         f"{self.area if self.area else '1/3'}*{self.density/1000}*0.03"]
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
