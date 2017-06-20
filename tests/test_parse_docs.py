def test_parse_doc(html_parser, html_document):
    html_parser.parse(html_document.read_text())

