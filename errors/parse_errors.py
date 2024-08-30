class ValidError(Exception):
    pass


class ParseError(Exception):
    pass


class ParseSpecieError(ParseError):
    pass


class UnknownSpecie(ParseSpecieError):
    pass


class ParseQuantityError(ParseError):
    pass


class ParseDiameterError(ParseError):
    pass


class ParseHeightError(ParseError):
    pass
