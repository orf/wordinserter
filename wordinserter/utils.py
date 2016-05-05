from wordinserter.operations import Text, ChildlessOperation


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


def pprint(tokens, indent=0):
    pad = '\t' * indent
    for token in tokens:
        if isinstance(token, Text):
            print(pad + '\t' + repr(token.text))
        elif isinstance(token, ChildlessOperation):
            print(pad + token.__class__.__name__)
        else:
            print(pad + token.__class__.__name__)
            pprint(list(token), indent + 1)
            print(pad + '/' + token.__class__.__name__)
