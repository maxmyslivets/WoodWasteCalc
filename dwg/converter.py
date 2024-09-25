from pathlib import Path
import subprocess
import shutil


def check_converter_in_system() -> bool:

    converter_path = Path(r'C:\Windows\System32') / "ODAFileConverter.exe"
    return converter_path.is_file()


def install_converter() -> None:

    source_file = Path('dwg/') / "ODAFileConverter.exe"
    destination_file = Path(r'C:\Windows\System32') / "ODAFileConverter.exe"

    try:
        shutil.copy2(source_file, destination_file)
        print(f'Конвертер успешно скопирован в папку System32.')
    except PermissionError:
        print(f'Ошибка: недостаточно прав для копирования в папку System32.')
    except Exception as e:
        print(f'Ошибка при копировании: {e}')


def dwg2dxf(input_dir: Path, output_dir: Path) -> None:

    converter_path = "ODAFileConverter"
    # Output version: ACAD9, ACAD10, ACAD12, ACAD14, ACAD2000, ACAD2004, ACAD2007, ACAD20010, ACAD2013, ACAD2018
    out_ver = "ACAD2018"
    # Output file type: DWG, DXF, DXB
    out_format = "DXF"
    # Recurse Input Folder: 0, 1
    recursive = "0"
    # Audit each file: 0, 1
    audit = "1"
    # (Optional) Input files filter: *.DWG, *.DXF
    input_filter = "*.DWG"

    # Command to run
    cmd = [converter_path, input_dir.absolute(), output_dir.absolute(), out_ver, out_format, recursive, audit,
           input_filter]

    # Run
    subprocess.run(cmd, shell=True)
