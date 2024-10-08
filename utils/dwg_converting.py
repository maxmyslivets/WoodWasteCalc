from pathlib import Path
import subprocess

from config.config import Config


def dwg2dxf() -> None:

    config = Config()
    input_dir = config.directories.dwg_directory.absolute()
    output_dir = config.directories.dxf_directory.absolute()
    converter_path = config.directories.converter_path.absolute()

    invalid_path = []
    for path in [input_dir, output_dir, converter_path]:
        if isinstance(path, Path) and ' ' in str(path):
            invalid_path.append(str(path))
    if invalid_path:
        print(f"Invalid paths: {invalid_path}\nПереместите папку с программой в директорию без пробелов в именах папок.")
        return

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
    cmd = [converter_path, input_dir, output_dir, out_ver, out_format, recursive, audit, input_filter]

    # Run
    subprocess.run(cmd, shell=True)     # TODO: Дождаться завершения конвертации
