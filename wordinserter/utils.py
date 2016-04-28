
class _NotFound(object):
    pass


class CombinedConstants(object):
    def __init__(self, *constants):
        self.constants = constants

    def __getattr__(self, item):
        for c in self.constants:
            result = getattr(c, item, _NotFound)
            if result is not _NotFound:
                return result

        raise AttributeError("No constant with the name {0} found".format(item))