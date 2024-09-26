import json
from pathlib import Path

from parsing.config import Config
from parsing.parse_xls import XLSParser, RawWood
from parsing.species_db import read_species_from_json


config = Config()


class ValidXLS:
    def __init__(self, filepath: Path) -> None:
        self.file = filepath
        self.shrub_species, self.wood_species = read_species_from_json(
            config.taxation_characteristics.species_json_path)
        with (open(config.taxation_characteristics.density_json_path, 'r', encoding='utf-8') as file):
            data = json.load(file)
            group_of_wood = data.get('group_of_wood', {})
            self.density_species = []
            for _, species in group_of_wood.items():
                self.density_species.extend(species)

    def check_specie(self, specie: str) -> bool:
        return False if specie not in self.shrub_species and specie not in self.wood_species else True

    def check_density(self, specie: str) -> bool:
        return False if specie not in self.density_species else True

    def check_valid(self) -> bool:

        is_valid = True

        table = XLSParser().parse(self.file)
        woods = []
        for row in table:
            woods.append(RawWood(*row))

        # обнаружение неизвестных пород
        unknown_specie = []
        for wood in woods:
            specie = wood._name.split()[0].lower()
            if not self.check_specie(specie) and specie not in unknown_specie:
                unknown_specie.append(specie)
        if len(unknown_specie) != 0:
            print(f"В файле `{self.file.name}` обнаружены неизвестные породы: {unknown_specie}")
            is_valid = False

        # обнаружение пород с неизвестной плотностью древесины
        unknown_density = []
        for wood in woods:
            specie = wood._name.split()[0].lower()
            if not self.check_density(specie) and specie not in unknown_density:
                unknown_density.append(specie)
        if len(unknown_density) != 0:
            print(f"В файле `{self.file.name}` обнаружены породы с неизвестной плотностью древесины: {unknown_density}")
            is_valid = False

        return is_valid
