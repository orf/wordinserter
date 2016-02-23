from . import BaseParser, ParseException

from ..operations import Paragraph, Bold, Italic, UnderLine, Text,\
    CodeBlock, Group, IgnoredOperation, Style, Image, HyperLink, BulletList,\
    NumberedList, ListElement, BaseList, Table, TableRow, TableCell, TableHead, Format, InlineCode

from .html import HTMLParser


class MarkdownParser(BaseParser):
    def __init__(self):
        raise RuntimeError("Markdown is not supported at this time. Blame the CommonMark Library")
        self.html_parser = HTMLParser()

    def parse(self, content):
        p = CommonMark.Parser()
        ast = p.parse(content)

        returner = []

        import sys
        sys.stdout.flush()

        for obj in ast.children:
            result = self._parse_node(obj)

            if result is None:
                continue

            returner.append(result)

        return returner

    def _parse_node(self, node):

        children = node.children + node.inline_content

        if node.t in ("ATXHeader", "SetextHeader"):
            obj = Style(name="Heading {0}".format(node.level))
        elif node.t == "Paragraph":
            obj = Paragraph()
        elif node.t == "Str":
            obj = Text(text=node.c)
        elif node.t in ("Softbreak", "Hardbreak"):
            obj = Text(text=" ")
        elif node.t == "Emph":
            obj = Italic()
        elif node.t == "Strong":
            obj = Bold()
        elif node.t == "Image":
            caption = node.label[0].c if len(node.label) else ""
            obj = Image(location=node.destination, caption=caption)
        elif node.t == "HtmlBlock":
            # Special case. Parse the HTML into instructions
            instructions = self.html_parser.parse(node.string_content)

            if not instructions:
                obj = IgnoredOperation()
            elif len(instructions) == 1:
                # Only contains one instruction, carry on as normal
                return instructions[0]
            else:
                # Lots of instructions. Return a group
                return Group(instructions)
        elif node.t == "List":
            obj = BulletList() if node.list_data["type"] == "Bullet" else NumberedList()
        elif node.t == "ListItem":
            obj = ListElement()
        elif node.t == "Link":
            obj = HyperLink(location=node.destination)
            children.extend(node.label)
        elif node.t in ("ReferenceDef", "HorizontalRule"):
            # ToDo: handle markdown references
            obj = IgnoredOperation()
        elif node.t == "Code":
            # Need an inline code object
            obj = InlineCode([Text(text=node.c)])
        elif node.t == "IndentedCode":
            obj = CodeBlock([Text(text=node.string_content)])
        else:
            CommonMark.dumpAST(node)
            raise ParseException("Cannot process node type {0}".format(node.t))

        if isinstance(node.c, list):
            children = children + node.c

        for child in children:
            result = self._parse_node(child)

            if result is None:
                continue

            obj.add_child(result)

        obj.set_source(node)
        return obj
