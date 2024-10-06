# Wood Waste Calc

Программа для расчёта массы древесных отходов при вырубке

* Конвертация файлов dwg в dxf
* Экспорт табличных значений из DWG в XLS
* Расчет массы отходов в выходные файлы XLS
* Вычисление суммарной массы отходов

## Настройка файла конфигурации `config.ini`

#### DIRECTORIES

```ini
# Относительный путь до директории с входными DWG файлами
dwg_directory = data/dwg
# Относительный путь до рабочей директории с DXF файлами
dxf_directory = data/dxf
# Относительный путь до рабочей директории с XLS файлами
xls_directory = data/xls
# Относительный путь до директории с выходными файлами
out_directory = data/out
```

#### TAXATION_CHARACTERISTICS

```ini
# Относительный путь до файла `species.json` с породами деревьев и кустарников
species_json_path = config/species.json
# Относительный путь до файла `volumes.xslx` с таблицей зависимости объема ствола от диаметра и высоты
volumes_xls_path = config/volumes.xlsx
# Относительный путь до файла `density.json` с плотностями пород
density_json_path = config/density.json
```

#### TABLE_STRUCTURE

```ini
# Литера колонки номеров
number = A
# Литера колонки наименований
specie = B
# Литера колонки количества
quality = C
# Литера колонки высот
height = D
# Литера колонки диаметров
diameter = E
```

## Руководство по использованию программы в режиме терминала

В файле конфигурации `config.ini` настроить режим работы на терминал.
```ini
[GUI]
gui = False
```

Порядок работы:

1. Положить DWG файлы в директорию для входных DWG файлов
2. Произвести конвертацию dwg файлов в dxf
3. Извлечь таблицы из dxf файлов в xls
4. Рассчитать массы отходов
5. Рассчитать итоговый файл
6. Забрать итоговый файл
7. Очистить рабочие директории


### Вариации входных данных
#### Номер

```text
1
2,3
4,5,6
6-7
9-14,15,17
15.2
```
#### Порода

В наименование породы берется только первое слово.

```text
Береза пушистая
Ива ломкая (54 ствола)
Ольха серая (6 стволов)
Береза пушистая (поросль)
сирень (поросль)
```

#### Количество

Столбец количества имеет либо число деревьев, либо площадь в метрах квадратных, либо длину в метрах погонных.

```text
2
172 м2
13м2
2м.п.
45 м.п.
```

#### Диаметр

```text
4 - один ствол диаметром 4
4,4 - два ствола с диаметрами 4
2х4 - 4 ствола с диаметрами по 2
3х7,11х4, 40х2 - 7 стволов по 3, 4 ствола по 11, 2 ствола по 40
30х4, 8х3,2,1 - 4 ствола по 30, 3 ствола по 8, 1 ствол диаметром 2, и 1 ствол диаметром 1
"-" - диаметр=2, кустарник
"1,5" - одно вещественное число, в случае, когда кол-во = 1 или площадь
```
Литера `х` может быть как на кириллице, так и на латинице.

#### Высота

```text
4 - один ствол высотой 4
4,4 - два ствола с высотой 4
2х4 - 4 ствола с высотами по 2
3х7,11х4, 40х2 - 7 стволов по 3, 4 ствола по 11, 2 ствола по 40
30х4, 8х3,2,1 - 4 ствола по 30, 3 ствола по 8, 1 ствол высотой 2, и 1 ствол высотой 1
"свыше 1м" - высота=2, кустарник
"от 1м" - высота=2, кустарник
"ниже 1м" - высота=1, кустарник
"до 1м" - высота=1, кустарник
"1,5" - одно вещественное число, в случае, когда кол-во = 1 или площадь
```
Литера `х` может быть как на кириллице, так и на латинице.