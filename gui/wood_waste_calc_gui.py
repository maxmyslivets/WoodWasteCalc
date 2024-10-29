import os
import shutil
import tempfile
import traceback
import webbrowser
from pathlib import Path

from calculation.waste import WasteCalculation
from gui.ui_main_window import Ui_MainWindow
from PySide6.QtWidgets import QApplication, QMainWindow, QListWidget, QListWidgetItem
from PySide6 import QtCore

from config.config import Config
from parsing.parse_xls import XLSWoodParser, RawWood
from validation.validation import ValidXLS
from utils.dwg_converting import dwg2dxf
from utils.table_extraction import table_extraction_from_dxf
from utils.xls_connection import xls_connection


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self) -> None:
        super(MainWindow, self).__init__()
        self.setupUi(self)

        self.label_path_to_dwg.setText(str(Config().directories.dwg_directory))
        self.label_path_to_dxf.setText(str(Config().directories.dxf_directory))
        self.label_path_to_xls.setText(str(Config().directories.xls_directory))
        self.label_path_to_out.setText(str(Config().directories.out_directory))

        self.button_setting.clicked.connect(self.open_settings)
        self.button_specie.clicked.connect(self.open_specie)
        self.button_density.clicked.connect(self.open_density)

        self.button_reload.clicked.connect(self.update_list_file)
        self.update_list_file()

        self.button_select_all_dwg.clicked.connect(self.select_all_dwg)
        self.button_select_all_dxf.clicked.connect(self.select_all_dxf)
        self.button_select_all_xls.clicked.connect(self.select_all_xls)
        self.button_select_all_out.clicked.connect(self.select_all_out)

        self.button_delete_dwg.clicked.connect(self.delete_selected_files_from_dwg)
        self.button_delete_dxf.clicked.connect(self.delete_selected_files_from_dxf)
        self.button_delete_xls.clicked.connect(self.delete_selected_files_from_xls)
        self.button_delete_out.clicked.connect(self.delete_selected_files_from_out)

        self.button_convert_dwg2dxf.clicked.connect(self.converting)
        self.button_extraction.clicked.connect(self.extraction)
        self.button_calculation.clicked.connect(self.calculation_waste)
        self.button_get_summary.clicked.connect(self.get_summary)

    @staticmethod
    def open_settings() -> None:
        webbrowser.open('config.ini')

    @staticmethod
    def open_specie() -> None:
        webbrowser.open(Config().taxation_characteristics.species_json_path)

    @staticmethod
    def open_density() -> None:
        webbrowser.open(Config().taxation_characteristics.density_json_path)

    def log(self, text) -> None:
        self.logout.append(str(text))

    def update_list_file(self) -> None:
        for list_file_widget, directory, suffix in [
            (self.list_files_dwg, Config().directories.dwg_directory, ".dwg"),
            (self.list_files_dxf, Config().directories.dxf_directory, ".dxf"),
            (self.list_files_xls, Config().directories.xls_directory, ".xlsx"),
            (self.list_files_out, Config().directories.out_directory, ".xlsx")
        ]:
            list_file_widget.clear()
            for file in [file for file in directory.iterdir() if file.is_file()]:
                if file.suffix != suffix:
                    continue
                item = QListWidgetItem(file.name)
                item.setCheckState(QtCore.Qt.CheckState.Checked)
                list_file_widget.addItem(item)

    @staticmethod
    def select_all(list_files: QListWidget) -> None:
        items = [list_files.item(i) for i in range(list_files.count())]
        if False in [item.checkState() == QtCore.Qt.CheckState.Checked for item in items]:
            for item in items:
                item.setCheckState(QtCore.Qt.CheckState.Checked)
        else:
            for item in items:
                item.setCheckState(QtCore.Qt.CheckState.Unchecked)

    def select_all_dwg(self) -> None:
        self.select_all(self.list_files_dwg)

    def select_all_dxf(self) -> None:
        self.select_all(self.list_files_dxf)

    def select_all_xls(self) -> None:
        self.select_all(self.list_files_xls)

    def select_all_out(self) -> None:
        self.select_all(self.list_files_out)

    def delete_selected_files(self, list_files: QListWidget, directory: Path) -> None:
        items = [list_files.item(i) for i in range(list_files.count())]
        for file in [directory / item.text() for item in items if item.checkState() == QtCore.Qt.CheckState.Checked]:
            if file.is_file():
                self.log(f"Удаление файла: `{file}`")
                os.remove(file)
        self.update_list_file()

    def delete_selected_files_from_dwg(self) -> None:
        self.delete_selected_files(self.list_files_dwg, Config().directories.dwg_directory)

    def delete_selected_files_from_dxf(self) -> None:
        self.delete_selected_files(self.list_files_dxf, Config().directories.dxf_directory)

    def delete_selected_files_from_xls(self) -> None:
        self.delete_selected_files(self.list_files_xls, Config().directories.xls_directory)

    def delete_selected_files_from_out(self) -> None:
        self.delete_selected_files(self.list_files_out, Config().directories.out_directory)

    def converting(self) -> None:
        selected_files = []
        items = [self.list_files_dwg.item(i) for i in range(self.list_files_dwg.count())]
        for file in [file for file in Config().directories.dwg_directory.iterdir()
                     if file.is_file() and file.suffix == ".dwg"]:
            for item in items:
                if item.checkState() == QtCore.Qt.CheckState.Checked and item.text() == file.name:
                    selected_files.append(file)
        temp_directory = Path(tempfile.gettempdir()) / "wood_waste_temp"
        if not temp_directory.exists():
            self.log(f"Создание временной директории `{temp_directory}`")
            temp_directory.mkdir()
        self.log(temp_directory)
        for file in temp_directory.iterdir():
            self.log(f"Удаление временного файла: `{file}`")
            os.remove(file)
        for file in selected_files:
            self.log(f"Копирование файла во временную директорию: `{file}`")
            shutil.copy(file, temp_directory)
        self.log("Конвертация файлов DWG в DXF...")
        try:
            dwg2dxf(temp_directory)
        except Exception as e:
            self.log(traceback.format_exc())

    def extraction(self) -> None:
        selected_files = []
        items = [self.list_files_dxf.item(i) for i in range(self.list_files_dxf.count())]
        for file in [file for file in Config().directories.dxf_directory.iterdir()
                     if file.is_file() and file.suffix == ".dxf"]:
            for item in items:
                if item.checkState() == QtCore.Qt.CheckState.Checked and item.text() == file.name:
                    selected_files.append(file)
        try:
            self.log(f"Извлечение таблицы из DFX файлов...")
            table_extraction_from_dxf(selected_files)
            self.log(f"DXF файлы успешно обработаны")
        except Exception as e:
            self.log(traceback.format_exc())
        self.update_list_file()

    def calculation_waste(self) -> None:
        selected_files = []
        items = [self.list_files_xls.item(i) for i in range(self.list_files_xls.count())]
        for file in [file for file in Config().directories.xls_directory.iterdir()
                     if file.is_file() and file.suffix == ".xlsx"]:
            for item in items:
                if item.checkState() == QtCore.Qt.CheckState.Checked and item.text() == file.name:
                    selected_files.append(file)
        try:

            xls_files_valid = selected_files.copy()

            # валидация таблиц
            for file in selected_files:
                valid_xls = ValidXLS(file, self.log)
                is_valid = valid_xls.check_valid()
                if not is_valid:
                    self.log(f"Файл `{file.name}` имеет нераспознанные данные и не будет обработан")
                    xls_files_valid.remove(file)

            # обработка таблиц
            for file in xls_files_valid:
                try:
                    raw_woods = XLSWoodParser().parse(file)
                except Exception as e:
                    self.log(f"`{file.name}` - Ошибка чтения данных из таблицы. {e}")
                    continue
                woods = []
                try:
                    for wood_list in raw_woods:
                        raw_wood = RawWood(*wood_list)
                        for wood in raw_wood.parse():
                            woods.append(wood)
                except Exception as e:
                    self.log(f"`{file.name}` - Ошибка извлечения данных из таблицы. {e}")
                    continue

                result_wood = [
                    [["номер", "порода", "количество", "диаметр", "высота", "объем", "плотность", "стволы", "сучья",
                      "пни"]], ]
                result_shrub = [[["номер", "порода", "количество", "диаметр", "высота", "объем", "площадь", "плотность",
                                  "сучья", "пни"]], ]
                for wood in woods:
                    w = WasteCalculation(wood.name, wood.number, wood.specie, wood.is_shrub, wood.trunks, wood.area)
                    w.export_preparation()
                    if w.is_shrub:
                        result_shrub.append(w.data)
                    else:
                        result_wood.append(w.data)

                if len(result_wood) > 1:
                    WasteCalculation.export_to_xls(result_wood, Config().directories.out_directory / f"{file.stem}_out_wood{file.suffix}")
                    self.log(f"Успешно `" + str(
                        Config().directories.out_directory / f"{file.stem}_out_wood{file.suffix}") + "`")
                if len(result_shrub) > 1:
                    WasteCalculation.export_to_xls(result_shrub, Config().directories.out_directory / f"{file.stem}_out_shrub{file.suffix}")
                    self.log(f"Успешно `" + str(
                        Config().directories.out_directory / f"{file.stem}_out_shrub{file.suffix}") + "`")
            self.log("Обработка XLSX файлов завершена")

        except Exception as e:
            self.log(traceback.format_exc())
        self.update_list_file()

    def get_summary(self) -> None:
        try:
            xls_connection()
            self.log("Суммарный файл успешно получен")
        except Exception as e:
            self.log(traceback.format_exc())
        self.update_list_file()


def main():
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
