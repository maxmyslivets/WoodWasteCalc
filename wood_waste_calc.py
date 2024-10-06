from config.config import Config

GUI = Config().gui.gui


if __name__ == '__main__':

    if not GUI:

        import os

        from utils.convert_dwg2dxf import convert_dwg_to_dxf
        from utils.table_extraction import table_extraction_from_dxf
        from not_gui.wood_waste_calc_terminal import main
        from utils.xls_connection import xls_connection

        MENU = {
            '1': 'Произвести конвертацию dwg файлов в dxf',
            '2': 'Извлечь таблицы из dxf файлов в xls',
            '3': 'Рассчитать массы отходов',
            '4': 'Рассчитать итоговый файл с суммарной массой отходов',
            '5': 'Очистить рабочие директории',
            '6': 'Выход'
        }

        while True:
            for key, value in MENU.items():
                print(f'{key}. {value}')
            menu = input(f"Выберите действие: ")
            if menu == '1':
                print("Выполняется конвертация ...")
                convert_dwg_to_dxf()
                print("Выполнено")
            elif menu == '2':
                print("Выполняется извлечение таблиц ...")
                table_extraction_from_dxf()
                print("Выполнено")
            elif menu == '3':
                print("Выполняется расчет массы отходов ...")
                main()
                print("Выполнено")
            elif menu == '4':
                print("Выполняется расчет итогового файла ...")
                xls_connection()
                print("Выполнено")
            elif menu == '5':
                directories = [
                    directory.absolute() for directory in [
                        Config().directories.dwg_directory,
                        Config().directories.dxf_directory,
                        Config().directories.xls_directory,
                        Config().directories.out_directory
                    ]
                ]
                for directory in directories:
                    for file in os.listdir(directory):
                        print(f"Remove {os.path.join(directory, file)}")
                        os.remove(os.path.join(directory, file))
            elif menu == '6':
                exit()
            else:
                print("Неверный ввод")
    else:
        from gui.wood_waste_calc_gui import main
