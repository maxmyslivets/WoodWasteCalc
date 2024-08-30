import os


def list_files_in_directory(directory_path):
    """
    Возвращает список файлов в указанной папке. Вложенные папки не сканируются.

    :param directory_path: Путь к папке
    :return: Список путей к файлам в указанной папке
    """
    files_with_paths = []

    # Получаем список всех элементов в папке
    for item in os.listdir(directory_path):
        item_path = os.path.join(directory_path, item)

        # Проверяем, является ли элемент файлом
        if os.path.isfile(item_path):
            files_with_paths.append(item_path)

    return files_with_paths
