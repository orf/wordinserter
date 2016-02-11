import pprint


class Operation(object):
    requires = set()
    optional = set()
    allowed_children = set()

    def __init__(self, children=None, **kwargs):
        self.parent = None
        self.children = children or []
        self.args = []
        self.format = None

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

    def add_child(self, child):
        if self.allowed_children and child.__class__.__name__ not in self.allowed_children:
            raise RuntimeError("Child {0} is not allowed".format(child.__class__.__name__))
        self.children.append(child)

    def add_children(self, children):
        for child in children:
            self.add_child(child)

    def replace_child(self, child, new_child):
        idx = self.children.index(child)
        self.children[idx] = new_child

    def has_child(self, child_class):
        return any(isinstance(c, child_class) for c in self.children)

    def __repr__(self):
        return "<{0}: {1}>".format(self.__class__.__name__, self.children)

    def set_parent(self, parent):
        self.parent = parent

    def set_parents(self, parent=None):
        self.set_parent(parent)

        for child in self.children:
            child.set_parents(self)


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
    pass


class InlineCode(Operation):
    pass


class LineBreak(ChildlessOperation):
    pass


class Format(Operation):
    optional = {
        "style",
        "font_size", "font_color"
    }

    @staticmethod
    def rgbstring_to_wdcolor(value):
        """
        Transform a string like rgb(199,12,15) into a wdColor format used by word
        :param value: A string like rgb(int,int,int)
        :return: An integer representation that Word understands
        """
        left, right = value.find("("), value.find(")")
        values = value[left+1:right].split(",")
        rgblist = [v.strip() for v in values]
        return int(rgblist[0]) + 0x100 * int(rgblist[1]) + 0x10000 * int(rgblist[2])

    @staticmethod
    def pixels_to_points(pixels):
        """
        Transform a pixel string into points (used by word).

        :param pixels: string optionally ending in px
        :return: an integer point representation
        """
        if isinstance(pixels, str):
            if pixels.endswith("px"):
                pixels = pixels[:-2]
            pixels = int(pixels)

        return pixels * 0.75


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

    @property
    def dimensions(self):
        """
        Returns row, column counts
        """
        if not self.children:
            return 0, 0
        # This assumes that the number of columns is uniform. Should be improved.
        return len(self.children), len(self.children[0].children)


class TableHead(IgnoredOperation):
    allowed_children = {"TableRow"}


class TableBody(IgnoredOperation):
    allowed_children = {"TableRow"}


class TableRow(Operation):
    allowed_children = {"TableCell"}


class TableCell(Operation):
    pass
