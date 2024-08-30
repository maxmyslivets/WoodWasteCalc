import sys

from parsing.config import Config
from parsing.directory import list_files_in_directory
from parsing.parse_xls import XLSParser, RawWood
from wood_objects.wood import Wood, Trunk

config = Config()


def parse() -> list[Wood]:

    # Чтение файлов таблиц
    files_for_parse = list_files_in_directory(config.settings.parse_directory)
    raw_woods = []
    for file in files_for_parse:
        wood_of_file = XLSParser().parse(file)
        raw_woods.extend(wood_of_file)

    woods = []
    # Парсинг таблиц
    for wood_list in raw_woods:
        raw_wood = RawWood(*wood_list)
        try:
            if raw_wood.is_valid():
                for wood in raw_wood.parse():
                    woods.append(wood)
        except Exception as e:
            print(e)

    # TODO: Обработка непрочитанных деревьев
    if input("Имеются ошибки в некоторых исходных данных. Продолжать (пустой ввода, если да)?") != "":
        sys.exit(0)

    for wood in woods:
        print(wood)

    return woods


def main() -> None:
    source_woods = parse()
