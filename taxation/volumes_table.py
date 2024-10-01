import string
from openpyxl import load_workbook

from config.config import Config

config = Config()


class Volume:
    def __init__(self):
        self.filepath = config.taxation_characteristics.volumes_xls_path
        self.columns = list(string.ascii_uppercase[1:40 + 1]) + ['A' + i for i in string.ascii_uppercase[:15]]
        self.table = []

        volume_wb = load_workbook(self.filepath)
        volume_ws = volume_wb.active
        for row in range(2, 35 + 2):
            table_row = []
            for col in self.columns:
                table_row.append(volume_ws[f'{col}{row}'].value)
            self.table.append(table_row)

    def get_volume(self, diameter: float, height: float) -> float:
        """
        Значение объёма ствола по срединному диаметру и длине.
        Сергеев П.Н. Лесная таксация, Приложение 2, стр.270-273.
        :param diameter: диаметр ствола, м
        :param height: высота ствола, см
        :return: объём ствола, м3
        """
        if diameter < 1:
            diameter = 1
        if diameter > 40:
            diameter = 40   # TODO: вывести зависимость для предсказания объема
        if height < 1:
            height = 1
        if height > 35:
            height = 35   # TODO: вывести зависимость для предсказания объема
        return self.table[round(height)-1][round(diameter)-1]
