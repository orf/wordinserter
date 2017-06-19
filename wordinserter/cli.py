"""wordinserter

Usage:
   wordinserter <path> [--debug] [--css=<path>] [--style=<literal>]

Options:
    --debug             Enable wordinserter debug output (warning: quite verbose)
    --css=<path>        Path to a CSS file to include whilst rendering
    --style=<literal>   Literal CSS to include whilst rendering
"""

from wordinserter import parse, insert
import pathlib
import sys
from contexttimer import Timer
from docopt import docopt


def get_file_contents(path):
    path = pathlib.Path(path)
    if not path.exists() and not path.is_file():
        print('{0} does not exist or is not a file'.format(path), file=sys.stderr)
        sys.exit(1)

    return path.read_text()


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

    with Timer(factor=1000) as t:
        parsed = parse(text, stylesheets=css)

    print('Parsed in {0:f} ms'.format(t.elapsed))

    from comtypes.client import CreateObject

    with Timer(factor=1000) as t:
        word = CreateObject("Word.Application")
        doc = word.Documents.Add()

    print('Opened word in {0:f} ms'.format(t.elapsed))

    word.Visible = True

    from comtypes.gen import Word as constants

    with Timer(factor=1000) as t:
        insert(parsed, document=doc, constants=constants, debug=arguments['--debug'])

    print('Inserted in {0:f} ms'.format(t.elapsed))

if __name__ == '__main__':
    run()
