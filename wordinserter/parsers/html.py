from . import BaseParser
from ..operations import Paragraph, Bold, Italic, UnderLine, Text,\
    CodeBlock, Group, IgnoredOperation, Style, Image, HyperLink, BulletList,\
    NumberedList, ListElement, BaseList, Table, TableRow, TableCell, TableHeading, Format
import bs4
from functools import partial
import cssutils


class HTMLParser(BaseParser):
    def __init__(self):
        # Preserve whitespace as-is
        self.preserve_whitespace = {
            CodeBlock
        }
        # Strip whitespace but keep spaces between tags
        self.respect_whitespace = {
            Paragraph, Bold, Italic, UnderLine, TableCell, Style, HyperLink
        }

        # Ignore all whitespace
        self.ignore_whitespace = {
            Table, TableRow, NumberedList, BulletList, ListElement
        }

        self.mapping = {
            "p": Paragraph,
            "b": Bold,
            "strong": Bold,
            "i": Italic,
            "em": Italic,
            "u": UnderLine,
            "pre": CodeBlock,
            "div": Group,

            "h1": partial(Style, name="Heading 1"),
            "h2": partial(Style, name="Heading 2"),
            "h3": partial(Style, name="Heading 3"),
            "h4": partial(Style, name="Heading 4"),

            "ul": BulletList,
            "ol": NumberedList,
            "li": ListElement,

            "img": Image,
            "a": HyperLink,
            "html": Group,

            "table": Table,
            "tr": TableRow,
            "td": TableCell,
            "th": TableHeading,
        }

    def parse(self, content):
        parser = bs4.BeautifulSoup(content)

        tokens = []

        for element in parser.childGenerator():
            item = self.build_element(element)

            if item is None:
                continue

            tokens.append(item)
        #import pprint
        #pprint.pprint(tokens)
        return tokens

    def build_element(self, element, whitespace="ignore"):
        if isinstance(element, bs4.Comment):
            return None

        if isinstance(element, bs4.NavigableString):
            if element.isspace():
                if whitespace == "preserve":
                    return Text(text=str(element))
                elif whitespace == "ignore":
                    return None
                elif whitespace == "respect":
                    if isinstance(element.previous_sibling, bs4.NavigableString):
                        return None
                    return Text(text=" ")
            return Text(text=str(element))

        cls = self.mapping.get(element.name, IgnoredOperation)

        if cls is Image:
            cls = partial(Image,
                          height=element.attrs.get("height", None),
                          width=element.attrs.get("width", None),
                          caption=element.attrs.get("alt", None),
                          location=element.attrs["src"])
        elif cls is HyperLink:
            if "href" not in element.attrs:
                cls = IgnoredOperation
            else:
                cls = partial(HyperLink, location=element.attrs["href"])

        instance = cls()

        if cls in self.respect_whitespace:
            whitespace = "respect"
        elif cls in self.preserve_whitespace:
            whitespace = "preserve"
        elif cls in self.ignore_whitespace:
            whitespace = "ignore"

        children = list(element.childGenerator())

        for idx, child in enumerate(children):
            item = self.build_element(child, whitespace=whitespace)
            if item is None:
                continue

            if isinstance(item, Text):
                if len(children) != 1:
                    item = item.keep_some_whitespace()

            if isinstance(instance, BaseList) and not isinstance(item, ListElement):
                # Wrap the item in a ListElement
                item = ListElement(children=[item])

            if isinstance(item, IgnoredOperation):
                instance.add_children(item.children)
            else:
                instance.add_child(item)

        args = {}

        for attribute, value in element.attrs.items():
            if attribute == "class" and value:
                # ToDo: Handle multiple classes? Idk.
                args["style"] = value[0]
            elif attribute == "style":
                styles = cssutils.parseStyle(value)
                for style in styles:
                    if style.name == "font-size":
                        args["font_size"] = Format.pixels_to_points(style.value)
                    elif style.name == "color":
                        args["font_color"] = Format.rgbstring_to_wdcolor(style.value)

        if args:
            instance.format = Format(**args)

        if cls in (Paragraph,):
            # Respect it but trim it on the ends
            while instance.children and \
                    (isinstance(instance.children[0], Text) or isinstance(instance.children[-1], Text)):
                first, last = instance.children[0], instance.children[-1]
                if hasattr(first, "text") and first.text.isspace():
                    instance.children.remove(first)
                elif hasattr(last, "text") and last.text.isspace():
                    instance.children.remove(last)
                else:
                    break

        return instance