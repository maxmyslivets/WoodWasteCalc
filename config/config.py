import configparser
from pathlib import Path


class ConfigSection:
    def __init__(self, config, section):
        self.config = config
        self.section = section

    def get(self, key, default=None):
        return self.config.get(self.section, key, fallback=default)

    def get_list(self, key, delimiter=', '):
        value = self.get(key)
        return value.split(delimiter) if value else []

    def get_int(self, key, default=0):
        return self.config.getint(self.section, key, fallback=default)

    def get_bool(self, key, default=False):
        return self.config.getboolean(self.section, key, fallback=default)


class Directories(ConfigSection):

    @property
    def dwg_directory(self) -> Path:
        return Path(self.get('dwg_directory'))

    @property
    def dxf_directory(self) -> Path:
        return Path(self.get('dxf_directory'))

    @property
    def xls_directory(self) -> Path:
        return Path(self.get('xls_directory'))

    @property
    def out_directory(self) -> Path:
        return Path(self.get('out_directory'))

    @property
    def converter_path(self) -> Path:
        return Path(self.get('converter_path'))


class TaxationCharacteristics(ConfigSection):

    @property
    def species_json_path(self) -> Path:
        return Path(self.get('species_json_path'))

    @property
    def volumes_xls_path(self) -> Path:
        return Path(self.get('volumes_xls_path'))

    @property
    def density_json_path(self) -> Path:
        return Path(self.get('density_json_path'))


class TableStructure(ConfigSection):

    @property
    def number(self) -> str:
        return self.get('number')

    @property
    def specie(self) -> str:
        return self.get('specie')

    @property
    def quality(self) -> str:
        return self.get('quality')

    @property
    def height(self) -> str:
        return self.get('height')

    @property
    def diameter(self) -> str:
        return self.get('diameter')


class Config:
    def __init__(self, filename: Path = Path('config.ini')):
        self.config = configparser.ConfigParser()
        self.config.read(filename, encoding='utf-8')

    @property
    def directories(self):
        return Directories(self.config, 'DIRECTORIES')

    @property
    def taxation_characteristics(self):
        return TaxationCharacteristics(self.config, 'TAXATION_CHARACTERISTICS')

    @property
    def table_structure(self):
        return TableStructure(self.config, 'TABLE_STRUCTURE')
