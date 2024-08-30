class Trunk:
    """
    Класс для создания объекта Ствол
    """
    def __init__(self, diameter: int, height: int) -> None:
        self.diameter = diameter
        self.height = height

    def __repr__(self) -> str:
        return f"Диаметр {self.diameter} см, высота {self.height} м"


class Wood:
    """
    Класс для создания объекта Дерево
    """
    def __init__(self, name: str, number: str, specie: str, is_shrub: bool, trunks: list[Trunk],
                 area: float = None) -> None:
        self.name = name
        self.number = number
        self.specie = specie
        self.is_shrub = is_shrub
        self.trunks = trunks
        self.area = area

    def __repr__(self) -> str:
        area = f", {self.area} м2" if self.area else ""
        if len(self.trunks) > 1:
            wood_str = f"{self.number}, {self.name}, {'кустарник' if self.is_shrub else 'дерево'}, " \
                       f"{len(self.trunks)} стволов"
            trunks_str = "\n".join([f"Ствол {idx+1}: {trunk}" for idx, trunk in enumerate(self.trunks)])
            return f"{wood_str}\n{trunks_str}"
        else:
            return f"{self.number}, {self.name}, {'кустарник' if self.is_shrub else 'дерево'}, {self.trunks[0]}"+area