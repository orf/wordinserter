from collections import defaultdict

from . import BaseParser
from ..operations import Paragraph, Bold, Italic, UnderLine, Text,\
    CodeBlock, Group, IgnoredOperation, Style, Image, HyperLink, BulletList,\
    NumberedList, ListElement, BaseList, Table, TableRow, TableCell, TableHead, TableBody, Format, Footnote, Span, \
    LineBreak
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
            Bold, Italic, UnderLine, Style, HyperLink, Paragraph
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
            "span": Span,

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
            "thead": TableHead,
            "tbody": TableBody,
            "tr": TableRow,
            "td": TableCell,
            "th": TableCell,

            "footnote": Footnote,
            "br": LineBreak
        }

    def parse(self, content):
        parser = bs4.BeautifulSoup(content, "lxml")

        tokens = []

        for element in parser.childGenerator():
            item = self.build_element(element)

            if item is None:
                continue

            tokens.append(item)
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
                          height=int(element.attrs.get("height", 0)),
                          width=int(element.attrs.get("width", 0)),
                          caption=element.attrs.get("alt", None),
                          location=element.attrs["src"])
        elif cls is HyperLink:
            if "href" not in element.attrs:
                cls = IgnoredOperation
            else:
                cls = partial(HyperLink, location=element.attrs["href"])
        elif cls is TableCell:
            cls = partial(TableCell,
                          colspan=int(element.attrs.get("colspan", 1)),
                          rowspan=int(element.attrs.get("rowspan", 1)))
        elif cls is Table:
            cls = partial(Table, border=element.attrs.get("border", "1"))
        elif cls is CodeBlock:
            highlight = element.attrs.get("highlight")
            text = element.getText()
            cls = partial(CodeBlock, highlight=highlight, text=text)

        instance = cls(attributes=element.attrs)
        cls = instance.__class__

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

            if isinstance(item, Text) and not whitespace == 'preserve':
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
                # This selects the first one which isn't an empty string. We could handle multiple classes here somehow.
                vals = [v for v in value if v]
                if vals:
                    args["style"] = vals[0]
            elif attribute == "style":
                styles = cssutils.parseStyle(value)
                args["margins"] = defaultdict(str)

                for style in styles:
                    if style.name == "font-size":
                        args["font_size"] = style.value
                    elif style.name == "color":
                        args["font_color"] = style.value
                    elif style.name == "background-color":
                        args["background_color"] = style.value
                    elif style.name == "text-decoration":
                        args["text_decoration"] = style.value
                    elif style.name.startswith("margin-"):
                        args["margins"][style.name.replace("margin-", "")] = style.value
                    elif style.name == "vertical-align":
                        args["vertical_align"] = style.value
                    elif style.name == "text-align":
                        args['horizontal_align'] = style.value

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
                elif first is last:
                    # Only one child, strip da text
                    instance[0].text = instance[0].text.strip()
                    break
                else:
                    break

        instance.set_source(element)

        return instance
