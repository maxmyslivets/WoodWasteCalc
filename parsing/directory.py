from pathlib import Path


def get_files_in_directory(directory_path: Path):
    """
    Возвращает список файлов в указанной папке. Вложенные папки не сканируются.

    :param directory_path: Путь к папке
    :return: Список путей к файлам в указанной папке
    """

    return [file for file in directory_path.iterdir() if file.is_file()]
