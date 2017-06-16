"""

"""

from wordinserter import parse, insert
import pathlib
import sys
from contexttimer import Timer


def run():
    if len(sys.argv) == 1:
        print('{0} [path to html file]'.format(sys.argv[0]), file=sys.stderr)
        sys.exit(1)

    path = pathlib.Path(sys.argv[1])
    if not path.exists() and not path.is_file():
        print('{0} does not exist or is not a file'.format(path), file=sys.stderr)
        sys.exit(1)

    text = path.read_text()

    with Timer(factor=1000) as t:
        parsed = parse(text)

    print('Parsed in {0:f} ms'.format(t.elapsed))

    from comtypes.client import CreateObject

    with Timer(factor=1000) as t:
        word = CreateObject("Word.Application")
        doc = word.Documents.Add()

    print('Opened word in {0:f} ms'.format(t.elapsed))

    word.Visible = True

    from comtypes.gen import Word as constants

    with Timer(factor=1000) as t:
        insert(parsed, document=doc, constants=constants)

    print('Inserted in {0:f} ms'.format(t.elapsed))

if __name__ == '__main__':
    run()
