
class RenderData(object):
    pass


class Operation(object):
    requires = set()
    optional = set()
    allowed_children = set()

    def __init__(self, children=None, **kwargs):
        self.parent = None
        self.children = children or []
        self.args = []
        self.format = None
        self.attributes = kwargs.pop("attributes", {})
        self.render = RenderData()

        self.source = None

        for k in self.optional:
            setattr(self, k, None)

        for kwarg, val in kwargs.items():
            if kwarg not in self.requires and kwarg not in self.optional:
                raise RuntimeError("Unexpected argument {0}".format(kwarg))
            setattr(self, kwarg, val)

        if self.allowed_children:
            for child in self.children:
                if child.__class__.__name__ not in self.allowed_children:
                    raise RuntimeError("Child {0} is not allowed!".format(child.__class__.__name__))

    def set_source(self, source):
        self.source = source

    def add_child(self, child):
        if self.allowed_children and child.__class__.__name__ not in self.allowed_children:
            raise RuntimeError("Child {0} is not allowed".format(child.__class__.__name__))
        self.children.append(child)

    def add_children(self, children):
        for child in children:
            self.add_child(child)

    def replace_child(self, child, new_child):
        self.children[self.child_index(child)] = new_child

    def has_child(self, child_class):
        return any(isinstance(c, child_class) for c in self.children)

    def child_index(self, child):
        return self.children.index(child)

    @property
    def previous_sibling(self):
        idx = self.parent.child_index(self)
        if idx == 0:
            return None

        return self.parent[idx-1]

    @property
    def next_sibling(self):
        idx = self.parent.child_index(self)
        if idx == (len(self.parent) - 1):
            return None

        return self.parent[idx + 1]

    @property
    def has_children(self):
        return len(self.children) > 0

    def __repr__(self):
        if len(self.children) == 1:
            child_repr = repr(self.children[0])
        else:
            child_repr = repr(self.children)

        return "<{0}: {1}>".format(self.__class__.__name__, child_repr)

    def set_parent(self, parent):
        self.parent = parent

    def set_parents(self, parent=None):
        self.set_parent(parent)

        for child in self.children:
            child.set_parents(self)

    def __len__(self):
        return len(self.children)

    def __getitem__(self, item):
        return self.children[item]


class ChildlessOperation(Operation):
    def __init__(self, **kwargs):
        kwargs["children"] = []
        super().__init__(**kwargs)

    def __repr__(self):
        return "<{0}>".format(self.__class__.__name__)

    def has_child(self, child_class):
        return False

    def replace_child(self, child, new_child):
        raise RuntimeError("ChildlessOperation: Cannot replace child")


class IgnoredOperation(Operation):
    pass


class Group(Operation):
    pass


class Bold(Operation):
    pass


class Italic(Operation):
    pass


class UnderLine(Operation):
    pass


class Text(ChildlessOperation):
    requires = {"text"}

    def __repr__(self):
        if len(self.text) > 10:
            txt = self.text[:10] + "..."
        else:
            txt = self.text
        return "<Text '{0}' />".format(txt)

    def strip_whitespace(self):
        return Text(text=self.text.strip())

    def keep_some_whitespace(self):
        # Don't know what to call this method.
        # If self.text has whitespace around it then strip it, but keep one space
        txt = self.text
        if txt[0].isspace():
            txt = " " + txt.lstrip()
        if txt[-1].isspace():
            txt = txt.rstrip() + " "

        return Text(text=txt)


class Paragraph(Operation):
    pass


class BlockParagraph(Operation):
    """
    Same as Paragraph but doesn't add a newline at the end
    """


class CodeBlock(Operation):
    optional = {"highlight", "text"}

    def highlighted_operations(self):
        from pygments.lexers import get_lexer_by_name
        from pygments.util import ClassNotFound
        from pygments import highlight
        from pygments.formatters import HtmlFormatter
        from wordinserter import parse
        import warnings

        try:
            lexer = get_lexer_by_name(self.highlight)
        except ClassNotFound:
            warnings.warn("Lexer {0} not found, not highlighting".format(self.highlight))
            return None

        highlighted_code = highlight(self.text, lexer=lexer, formatter=HtmlFormatter(noclasses=True))
        return parse(highlighted_code, parser="html")


class InlineCode(Operation):
    pass


class LineBreak(ChildlessOperation):
    pass


class Span(Operation):
    pass


class Format(Operation):
    optional = {
        "style",
        "font_size",
        "font_color",
        "background_color",
        "text_decoration",
        "margins",
        "vertical_align",
        "horizontal_align"
    }

    def has_format(self):
        return any(getattr(self, name) for name in self.optional)

    def __repr__(self):
        return "<{0}: {1}>".format(self.__class__.__name__,
                                   {n: getattr(self, n) for n in self.optional if getattr(self, n) is not None})


class Style(Operation):
    requires = {"name"}


class Font(Operation):
    optional = {"size", "color"}


class Image(ChildlessOperation):
    requires = {"location"}
    optional = {"height", "width", "caption"}


class HyperLink(Operation):
    requires = {"location"}
    optional = {"label"}


class BaseList(Operation):
    allowed_children = {"ListElement"}
    pass


class BulletList(BaseList):
    pass


class NumberedList(BaseList):
    pass


class ListElement(Operation):
    pass


class Table(Operation):
    allowed_children = {"TableRow", "TableHead", "TableBody"}
    optional = {"border"}

    @property
    def dimensions(self):
        """
        Returns row, column counts
        """
        if not self.children:
            return 0, 0

        rows = len(self.children)
        columns = max(sum(child.colspan or 1 for child in row.children) for row in self.children)

        return rows, columns


class TableHead(IgnoredOperation):
    allowed_children = {"TableRow"}


class TableBody(IgnoredOperation):
    allowed_children = {"TableRow"}


class TableRow(Operation):
    allowed_children = {"TableCell"}


class TableCell(Operation):
    optional = {"colspan", "rowspan"}


class Footnote(ChildlessOperation):
    pass
