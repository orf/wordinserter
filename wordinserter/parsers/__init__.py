class ParseException(RuntimeError):
    pass


class BaseParser(object):
    def parse(self, content):
        raise NotImplementedError()


from .html import HTMLParser
from .markdown import MarkdownParser
