def test_parse_doc(html_parser, html_document):
    with html_document.open() as fd:
        html_parser.parse(fd.read())

