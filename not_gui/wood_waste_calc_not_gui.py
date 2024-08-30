from parsing.config import Config
from parsing.directory import list_files_in_directory
from parsing.parse_xls import XLSParser, RawWood

config = Config()


def main():

    # Чтение файлов таблиц
    files_for_parse = list_files_in_directory(config.settings.parse_directory)
    woods = []
    for file in files_for_parse:
        wood_of_file = XLSParser().parse(file)
        woods.extend(wood_of_file)

    # Парсинг таблиц
    for wood_list in woods:
        raw_wood = RawWood(*wood_list)
        try:
            if raw_wood.is_valid():
                for wood in raw_wood.parse():
                    print(wood)
        except Exception as e:
            print(e)


if __name__ == '__main__':
    main()
