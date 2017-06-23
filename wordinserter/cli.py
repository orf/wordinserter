"""wordinserter

Usage:
   wordinserter <path> [--debug] [--css=<path>] [--style=<literal>] [--save=<name>] [--close] [--hidden]

Options:
    --debug             Enable wordinserter debug output (warning: quite verbose)
    --css=<path>        Path to a CSS file to include whilst rendering
    --style=<literal>   Literal CSS to include whilst rendering
    --save=<name>       Save the document to this path. Format is defined by the extension (e.g output.pdf).
    --close             Close the word document after rendering
    --hidden            Hide the Word window while rendering
"""

import pathlib
import sys
import tempfile
import inspect
import os

from comtypes import gen
from contexttimer import Timer
from docopt import docopt

from comtypes.client import CreateObject
from wordinserter import insert, parse


def get_file_contents(path):
    path = pathlib.Path(path)
    if not path.exists() and not path.is_file():
        print('{0} does not exist or is not a file'.format(path), file=sys.stderr)
        exit(1)

    return path.read_text()


def save_as_image(document, save_as, constants):
    from wand.image import Image
    from wand.color import Color
    from wand.exceptions import DelegateError

    temp_directory = pathlib.Path(tempfile.mkdtemp())

    temp_pdf_name = temp_directory / (save_as.name + '.pdf')

    document.SaveAs2(
        FileName=str(temp_pdf_name.absolute()),
        FileFormat=constants.wdFormatPDF,
    )

    try:
        with Image(filename=str(temp_pdf_name), resolution=300) as pdf:
            pdf.trim()

            with Image(width=pdf.width, height=pdf.height, background=Color('white')) as png:
                png.composite(pdf, 0, 0)
                png.save(filename=str(save_as))
    except DelegateError:
        print('Error: You may need to install ghostscript as well', file=sys.stderr)
        exit(1)

# https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat.aspx

SAVE_FORMATS = {
    '.pdf': 'wdFormatPDF',
    '.xps': 'wdFormatXPS',
    '.rtf': 'wdFormatRTF',
    '.doc': 'wdFormatDocument',
    '.docx': 'wdFormatDocument',
    '.png': save_as_image,
    '.jpeg': save_as_image
}


def run():
    arguments = docopt(__doc__, version='0.1')

    if arguments['<path>'] == '-':
        text = sys.stdin.read()
    else:
        text = get_file_contents(arguments['<path>'])

    css = []

    if arguments['--css']:
        css.append(get_file_contents(arguments['--css']))

    if arguments['--style']:
        css.append(arguments['--style'])

    save_as = None
    if arguments['--save']:
        save_as = pathlib.Path(arguments['--save'])
        if save_as.suffix not in SAVE_FORMATS:
            print('Error: Cannot save in {0} format. Supported formats: {1}'.format(
                save_as.suffix,
                ', '.join(SAVE_FORMATS.keys())
            ), file=sys.stderr)
            exit(1)

        if save_as.exists():
            print('Error: Path {0} already exists. Not overwriting'.format(save_as), file=sys.stderr)
            exit(1)

    with Timer(factor=1000) as t:
        parsed = parse(text, stylesheets=css)

    print('Parsed in {0:f} ms'.format(t.elapsed))

    with Timer(factor=1000) as t:
        try:
            word = CreateObject("Word.Application")
        except AttributeError as e:
            gen_dir = inspect.getsourcefile(gen)

            print('****** There was an error opening word ******')
            print('This is a transient error that sometimes happens.')
            print('Remove all files (except __init__.py) from here:')
            print(os.path.dirname(gen_dir))
            print('Then retry the program')
            print('*********************************************')
            raise e
        doc = word.Documents.Add()

    print('Opened word in {0:f} ms'.format(t.elapsed))

    word.Visible = not arguments['--hidden']

    from comtypes.gen import Word as constants

    with Timer(factor=1000) as t:
        insert(parsed, document=doc, constants=constants, debug=arguments['--debug'])

    print('Inserted in {0:f} ms'.format(t.elapsed))

    if save_as:
        file_format_attr = SAVE_FORMATS[save_as.suffix]

        if callable(file_format_attr):
            file_format_attr(doc, save_as, constants)
        else:
            # https://msdn.microsoft.com/en-us/library/microsoft.office.tools.word.document.saveas2.aspx
            doc.SaveAs2(
                FileName=str(save_as.absolute()),
                FileFormat=getattr(constants, file_format_attr),
            )
        print('Saved to {0}'.format(save_as))

    if arguments['--close']:
        word.Quit(SaveChanges=constants.wdDoNotSaveChanges)


if __name__ == '__main__':
    run()
