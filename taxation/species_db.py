import json

from config.config import Config

config = Config()


class Species:
    def __init__(self):
        self.filepath = config.taxation_characteristics.species_json_path

    def get_species(self):
        """
        Читает списки shrub_species и wood_species из JSON файла.

        :return: Кортеж с двумя списками (shrub_species, wood_species)
        """
        try:
            with open(self.filepath, 'r', encoding='utf-8') as file:
                data = json.load(file)
                shrub_species = data.get('shrub', [])
                wood_species = data.get('wood', [])
                return shrub_species, wood_species
        except FileNotFoundError as e:
            # Возвращаем пустые списки, если файл не найден
            print(f"`Ошибка чтения species.json {e}")
        except json.JSONDecodeError as e:
            # Возвращаем пустые списки, если произошла ошибка при чтении JSON
            print(f"`Ошибка чтения species.json {e}")

    def add_species_to_json(self, new_shrub_species=None, new_wood_species=None):
        """
        Добавляет новые элементы в списки shrub_species и wood_species и сохраняет их в JSON файл.

        :param new_shrub_species: Список новых кустарников для добавления (по умолчанию None)
        :param new_wood_species: Список новых древесных пород для добавления (по умолчанию None)
        """
        new_shrub_species = new_shrub_species or []
        new_wood_species = new_wood_species or []

        # Читаем текущие данные из файла
        shrub_species, wood_species = self.get_species()

        # Добавляем новые элементы, избегая дублирования
        shrub_species.extend(species for species in new_shrub_species if species not in shrub_species)
        wood_species.extend(species for species in new_wood_species if species not in wood_species)

        # Сохраняем обновленные списки обратно в JSON файл
        with open(self.filepath, 'w', encoding='utf-8') as file:
            json.dump({'shrub': shrub_species, 'wood': wood_species}, file, ensure_ascii=False, indent=4)
