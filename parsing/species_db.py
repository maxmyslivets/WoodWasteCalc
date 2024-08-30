import json


def read_species_from_json(filename):
    """
    Читает списки shrub_species и wood_species из JSON файла.

    :param filename: Имя JSON файла
    :return: Кортеж с двумя списками (shrub_species, wood_species)
    """
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            data = json.load(file)
            shrub_species = data.get('shrub', [])
            wood_species = data.get('wood', [])
            return shrub_species, wood_species
    except FileNotFoundError:
        # Возвращаем пустые списки, если файл не найден
        return [], []
    except json.JSONDecodeError:
        # Возвращаем пустые списки, если произошла ошибка при чтении JSON
        return [], []


def add_species_to_json(filename, new_shrub_species=None, new_wood_species=None):
    """
    Добавляет новые элементы в списки shrub_species и wood_species и сохраняет их в JSON файл.

    :param filename: Имя JSON файла
    :param new_shrub_species: Список новых кустарников для добавления (по умолчанию None)
    :param new_wood_species: Список новых древесных пород для добавления (по умолчанию None)
    """
    new_shrub_species = new_shrub_species or []
    new_wood_species = new_wood_species or []

    # Читаем текущие данные из файла
    shrub_species, wood_species = read_species_from_json(filename)

    # Добавляем новые элементы, избегая дублирования
    shrub_species.extend(species for species in new_shrub_species if species not in shrub_species)
    wood_species.extend(species for species in new_wood_species if species not in wood_species)

    # Сохраняем обновленные списки обратно в JSON файл
    with open(filename, 'w', encoding='utf-8') as file:
        json.dump({'shrub': shrub_species, 'wood': wood_species}, file, ensure_ascii=False, indent=4)
