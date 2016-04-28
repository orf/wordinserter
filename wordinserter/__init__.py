
from .parsers import HTMLParser, MarkdownParser, BaseParser
from .renderers import COMRenderer, BaseRenderer
from .exceptions import InsertError
from .utils import CombinedConstants
import inspect

parsers = {
    "html": HTMLParser,
    "markdown": MarkdownParser
}

renderers = {
    "com": COMRenderer
}


def parse(text, parser=None, **kwargs):
    """
    Parse some given input into a list of operations to perform
    :param text: Text input
    :param parser: Either 'html' or 'markdown', or a class that inherits from BaseParser
    :return: A list of operations
    """
    if isinstance(parser, str) and parser not in parsers:
        raise RuntimeError("Format {0} not recognized".format(format))

    # If we have been given a string instead of a parsers class then lookup the class from the parsers dictionary
    if not inspect.isclass(parser):
        parser = parsers[parser]

    parser = parser()

    return parser.parse(text, **kwargs)


def insert(operations, renderer="com", **kwargs):
    """
    Render a list of operations to a word document using the specified renderer
    :param operations: A sequence of operations to execute
    :param renderer: Either a string (only 'COM' supported at this time) or a class that inherits from BaseRenderer
    :param kwargs: Keyword arguments to pass to the renderer
    """
    if isinstance(renderer, str) and renderer not in renderers:
        raise RuntimeError("Unknown renderer {0}".format(renderer))

    if not inspect.isclass(renderer):
        renderer = renderers[renderer]

    renderer = renderer(**kwargs)
    renderer.render(operations)


def print_operations(operations, indent_level=0):
    indent = "  " * indent_level if indent_level > 0 else ""

    for op in operations:
        print(indent + op.__class__.__name__)
        print_operations(op.children, indent_level + 1)
