from collections import defaultdict

from wordinserter.parsers.fixes import normalize_list_elements, correct_whitespace
from . import BaseParser
from ..operations import Paragraph, Bold, Italic, UnderLine, Text,\
    CodeBlock, Group, IgnoredOperation, Style, Image, HyperLink, BulletList,\
    NumberedList, ListElement, Table, TableRow, TableCell, TableHead, TableBody, Format, Footnote, Span, \
    LineBreak, Heading
import bs4
from functools import partial
import cssutils
import re

_COLLAPSE_REGEX = re.compile(r'\s+')


class HTMLParser(BaseParser):
    def __init__(self):
        self.mapping = {
            "p": Paragraph,
            "b": Bold,
            "strong": Bold,
            "i": Italic,
            "em": Italic,
            "u": UnderLine,
            "code": CodeBlock,
            "pre": CodeBlock,
            "div": Group,
            "span": Span,

            "h1": partial(Heading, level=1),
            "h2": partial(Heading, level=2),
            "h3": partial(Heading, level=3),
            "h4": partial(Heading, level=4),

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

    def parse(self, content, stylesheets=None):
        parser = bs4.BeautifulSoup(content, "lxml")

        if stylesheets:
            # Iterate through each stylesheet, and each rule within each sheet, and apply the relevant styles as
            # inline-styles.
            docs = (cssutils.parseString(css_content) for css_content in stylesheets if css_content)
            for doc in docs:
                for rule in (rule for rule in doc.cssRules if rule.typeString == 'STYLE_RULE'):
                    rule_styles = dict(rule.style)
                    for selector in rule.selectorList:
                        elements = parser.select(selector.selectorText)
                        for element in elements:
                            style = cssutils.parseStyle(element.attrs.get("style", ""))
                            element_style = dict(style)
                            element_style.update(rule_styles)
                            for key, value in element_style.items():
                                style[key] = value
                            element.attrs["style"] = style.getCssText(" ")

        tokens = []

        for element in parser.childGenerator():
            item = self.build_element(element)

            if item is None:
                continue

            tokens.append(item)

        tokens = Group(tokens)
        normalize_list_elements(tokens)

        tokens.set_parents()
        correct_whitespace(tokens)

        return tokens

    def build_element(self, element):
        if isinstance(element, bs4.Comment):
            return None

        if isinstance(element, bs4.NavigableString):
            return Text(text=str(element))

        cls = self.mapping.get(element.name, IgnoredOperation)

        if cls is Image:
            if not element.attrs.get("src", None):
                cls = IgnoredOperation
            else:
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
        children = list(element.childGenerator())

        for idx, child in enumerate(children):
            item = self.build_element(child)
            if item is None:
                continue

            if isinstance(item, IgnoredOperation):
                instance.add_children(item.children)
            elif not instance.is_child_allowed(item):
                continue
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
                    args["style"] = vals
            elif attribute == "style":
                styles = cssutils.parseStyle(value)

                for name in Format.NESTED_STYLES:
                    args[name] = defaultdict(str)

                for style in styles:
                    for nested_name in Format.NESTED_STYLES:
                        nested_name_with_dash = nested_name + "-"
                        if style.name.startswith(nested_name_with_dash):
                            args[nested_name][style.name.replace(nested_name_with_dash, "")] = style.value
                            break
                    else:
                        name = style.name.lower().replace("-", "_")
                        if name in Format.NESTED_STYLES:
                            # Not supported. Use explicit 'margin-right',
                            # 'margin-left' etc rather than just 'margin'.
                            continue
                        elif name in Format.FORMAT_ALIASES:
                            name = Format.FORMAT_ALIASES[name]

                        if name in Format.optional:
                            args[name] = style.value.strip()

        instance.format = Format(**args)
        instance.set_source(element)

        return instance
