import abc


class ParseException(RuntimeError):
    pass


class BaseParser(abc.ABC):
    @abc.abstractmethod
    def parse(self, content):
        raise NotImplementedError()


from .html import HTMLParser
from .markdown import MarkdownParser
