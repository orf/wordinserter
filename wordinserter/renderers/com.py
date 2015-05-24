from . import BaseRenderer, renders
from ..operations import Text, Bold, Italic, UnderLine, Paragraph, LineBreak, CodeBlock, Style, Image, HyperLink, \
    BulletList, NumberedList, ListElement, BaseList, Table, TableCell, TableRow, TableHeading, Format, InlineCode
import warnings


class COMRenderer(BaseRenderer):
    def __init__(self, document, constants, range=None, debug=False):
        self.word = document.Application
        self.document = document
        self.constants = constants

        if range is not None:
            range.Select()

        super().__init__(debug)

    @property
    def selection(self):
        return self.document.ActiveWindow.Selection

    def range(self, start=None, end=None):
        if not (start or end):
            raise RuntimeError("Start and End are both None!")

        if start is None:
            start = end
        elif end is None:
            end = start

        return self.document.Range(Start=start, End=end)

    @renders(Style)
    def style(self, op):
        old_style = self.selection.Style
        self.selection.Style = self.document.Styles(op.name)
        yield
        self.selection.TypeParagraph()
        #self.selection.Collapse(Direction=constants.wdCollapseEnd)
        #self.selection.Style = old_style

    @renders(Bold)
    def bold(self, op):
        self.selection.BoldRun()
        yield
        self.selection.BoldRun()

    @renders(Italic)
    def italic(self, op):
        self.selection.ItalicRun()
        yield
        self.selection.ItalicRun()

    @renders(UnderLine)
    def underline(self, op):
        self.selection.Font.Underline = self.constants.wdUnderlineSingle
        yield
        self.selection.Font.Underline = self.constants.wdUnderlineNone

    @renders(Text)
    def text(self, op):
        self.selection.TypeText(op.text)

    @renders(LineBreak)
    def linebreak(self, op):
        self.selection.TypeParagraph()

    @renders(Paragraph)
    def paragraph(self, op):
        previous_style = None
        if op.has_child(LineBreak):
            previous_style = self.selection.Style
            self.selection.Style = self.document.Styles("No Spacing")

        yield

        if previous_style is not None:
            self.selection.Style = previous_style

        should_do_newline = True

        if op.children and isinstance(op.children[-1], (BaseList, Image, Table)):
            should_do_newline = False

        if isinstance(op.parent, ListElement) and op.parent.children[-1] is op:
            # If our parent is a ListElement and our operation is the last one of it's children then we don't need to
            # add a newline.
            should_do_newline = False

        if should_do_newline:
            self.selection.TypeParagraph()

    @renders(InlineCode)
    def inline_code(self, op):
        previous_style = self.selection.Style
        previous_font_name, previous_font_size = self.selection.Font.Name, self.selection.Font.Size
        self.selection.Font.Name = "Courier New"
        self.selection.Font.Size = 7

        yield

        self.selection.Style = previous_style
        self.selection.Font.Name = previous_font_name
        self.selection.Font.Size = previous_font_size

    @renders(CodeBlock)
    def code_block(self, op):
        previous_style = self.selection.Style
        previous_font_name, previous_font_size = self.selection.Font.Name, self.selection.Font.Size
        self.selection.Style = self.document.Styles("No Spacing")
        self.selection.Font.Name = "Courier New"
        self.selection.Font.Size = 7

        yield

        self.selection.ParagraphFormat.LineSpacingRule = self.constants.wdLineSpace1pt5
        self.selection.TypeParagraph()
        self.selection.Style = previous_style
        self.selection.Font.Name = previous_font_name
        self.selection.Font.Size = previous_font_size

    @renders(Image)
    def image(self, op):
        image = self.selection.InlineShapes.AddPicture(
            FileName=op.location)

        if op.height:
            image.Height = op.height

        if op.width:
            image.Width = op.width

        if op.caption:
            self.selection.TypeParagraph()
            self.selection.Range.Style = self.document.Styles("caption")
            self.selection.TypeText(op.caption)

        self.selection.TypeParagraph()

    @renders(HyperLink)
    def hyperlink(self, op):
        start_range = self.selection.Range.End
        yield
        # Inserting a hyperlink that contains different styles can reset the style. IE:
        # Link<Bold<Text>> Text
        # Bold will turn bold off, but link will reset it meaning the second Text is bold.
        # Here we just reset the style after making the hyperlink.
        style = self.selection.Style
        rng = self.document.Range(Start=start_range, End=self.selection.Range.End)
        self.document.Hyperlinks.Add(Anchor=rng, Address=op.location)
        self.selection.Collapse(Direction=self.constants.wdCollapseEnd)
        self.selection.Style = style

    @renders(BulletList, NumberedList)
    def render_list(self, op):
        first_list = self.selection.Range.ListFormat.ListTemplate is None

        if first_list:
            self.selection.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate=self.word.ListGalleries(
                    self.constants.wdNumberGallery if isinstance(op, NumberedList) else self.constants.wdBulletGallery
                ).ListTemplates(1),
                ContinuePreviousList=False,
                DefaultListBehavior=self.constants.wdWord10ListBehavior
            )
        else:
            self.selection.Range.ListFormat.ListIndent()

        yield

        if first_list:
            self.selection.Range.ListFormat.RemoveNumbers(NumberType=self.constants.wdNumberParagraph)
            #self.selection.TypeParagraph()
        else:
            self.selection.Range.ListFormat.ListOutdent()

    @renders(ListElement)
    def list_element(self, op):
        yield
        if not isinstance(op.children[-1], BaseList):
            self.selection.TypeParagraph()

    @renders(Table)
    def table(self, op):
        table_range = self.selection.Range
        self.selection.TypeParagraph()
        end_range = self.selection.Range

        rows, columns = op.dimensions

        table = self.selection.Tables.Add(
            table_range,
            NumRows=rows,
            NumColumns=columns,
            AutoFitBehavior=self.constants.wdAutoFitFixed
        )
        table.Style = "Table Grid"
        table.AllowAutoFit = True

        for idx, child in enumerate(op.children):
            child._row = table.Rows(idx + 1)
        table.Select()
        yield
        table.Columns.AutoFit()
        end_range.Select()

    @renders(TableRow)
    def table_row(self, op):
        row_index = op.parent.children.index(op) + 1
        yield row_index,

    @renders(TableCell, TableHeading)
    def table_cell(self, op, row_index):
        cell_index = op.parent.children.index(op) + 1
        cell_range = self.selection.Tables(1).Rows(row_index).Cells(cell_index).Range
        cell_range.End -= 1
        cell_range.Select()
        yield

    @renders(Format)
    def format(self, op):
        start = self.selection.Start
        yield
        end = self.selection.End

        element_range = self.range(start, end)

        if op.style:
            try:
                element_range.Style = op.style
            except Exception:
                warnings.warn("Unable to apply style name '{0}'".format(op.style))

        if op.font_size:
            element_range.Font.Size = op.font_size

        if op.font_color:
            element_range.Font.Color = op.font_color