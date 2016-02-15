import sys
from wordinserter import parse, insert
import pathlib
from comtypes.client import CreateObject
from comtypes import CoInitialize


def main():
    html = pathlib.Path(sys.argv[1]).read_text()

    word = CreateObject("Word.Application")
    word.visible = True
    from comtypes.gen import Word as c
    doc = word.Documents.Add()

    operations = parse(html)
    insert(operations, document=doc, constants=c, debug=True)


if __name__ == '__main__':
    main()
