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
            Bold, Italic, UnderLine, Style, HyperLink, Paragraph, ListElement
        }

        # Ignore all whitespace
        self.ignore_whitespace = {
            Table, TableRow, NumberedList, BulletList, Span
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

        self.normalize_list_elements(tokens)
        return tokens

    def normalize_list_elements(self, tokens):
        for token in tokens:
            if isinstance(token, BaseList):
                self.normalize_list(token)
            else:
                self.normalize_list_elements(token.children)

    def normalize_list(self, op: BaseList):
        # If there are > 1 lists to move out then we need to insert it after previously moved ones,
        # instead of before. `moved` tracks this.
        children = list(op)

        for child in children:
            if isinstance(child, ListElement):
                moved = 0
                for element_child in child:
                    if isinstance(element_child, BaseList):
                        moved += 1
                        # Move the list outside of the ListElement
                        child_index = op.child_index(child)
                        op.insert_child(child_index + moved, element_child)
                        child.remove_child(element_child)
                        self.normalize_list(element_child)

    def build_element(self, element, whitespace="ignore"):
        if isinstance(element, bs4.Comment):
            return None

        if isinstance(element, bs4.NavigableString):
            if whitespace == "preserve":
                return Text(text=str(element))

            elif whitespace == "ignore":
                if element.isspace():
                    return None

                return Text(text=element.strip())

            elif whitespace == "respect":
                if element.isspace():
                    if isinstance(element.previous_sibling, bs4.NavigableString):
                        return None
                    return Text(text=" ")

                if element[0].isspace():
                    element = " " + element.lstrip()
                if element[-1].isspace():
                    element = element.rstrip() + " "

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
        elif cls is NumberedList:
            type = element.attrs.get("type")
            values = {
                "i": "roman-lowercase",
                "I": "roman-uppercase",

            }
            cls = partial(NumberedList, type=values.get(type))

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

            if isinstance(item, IgnoredOperation):
                instance.add_children(item.children)
            else:
                instance.add_child(item)

        if instance.requires_children and not instance.children:
            return None

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
                    if style.name.startswith("margin-"):
                        args["margins"][style.name.replace("margin-", "")] = style.value
                    else:
                        name = style.name.lower().replace("-", "_")
                        if name in Format.optional:
                            args[name] = style.value.strip()

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
