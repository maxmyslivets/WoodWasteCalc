import json
from typing import Literal

from errors.parse_errors import UnknownSpecie
from parsing.config import Config

config = Config()


class Density:
    def __init__(self):
        try:
            with open(config.settings.density_json_path, 'r', encoding='utf-8') as file:
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
        raise UnknownSpecie(f"Попытка получения плотности для неизвестной породы '{specie}'.")

    def add_specie_to_group(self, group: Literal["мягкие", "твердые", "лиственница"], specie: str) -> None:
        """
        Добавляет новый вид древесины в указанный список группы.
        :param group: группа плотности ("мягкие", "твердые" или "лиственница")
        :param specie: порода древесины
        :return: None
        """

        if group not in self.group_of_wood:
            raise ValueError(f"Группа '{group}' не найдена. Доступные группы: "
                             f"{', '.join(self.group_of_wood.keys())}")     # TODO: добавить Exception

        if specie in self.group_of_wood[group]:
            print(f"Порода '{specie}' уже присутствует в группе '{group}'.")    # TODO: добавить Exception
            return

        self.group_of_wood[group].append(specie)
        try:
            with open(config.settings.density_json_path, 'w', encoding='utf-8') as file:
                json.dump({
                    'group_of_wood': self.group_of_wood,
                    'density_dict': self.density_dict
                }, file, ensure_ascii=False, indent=4)
        except Exception as e:
            print("Ошибка записи в файл density.json", e)   # TODO: добавить Exception
