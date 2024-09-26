import json
from typing import Literal

from errors.parse_errors import UnknownSpecie
from parsing.config import Config

config = Config()


class Density:
    def __init__(self):
        try:
            with open(config.taxation_characteristics.density_json_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
                self.group_of_wood = data.get('group_of_wood', {})
                self.density_dict = data.get('density_dict', {})
        except Exception as e:
            print("Ошибка чтения файла density.json", e)    # TODO: добавить Exception

    def get_density(self, specie: str) -> int:
        """
        Определяет плотность породы.
        :param specie: порода
        :return: плотность кг/м3
        """
        for group, species_list in self.group_of_wood.items():
            if specie in species_list:
                return self.density_dict.get(group)
