import pathlib

import pytest

from wordinserter import parsers

docs = pathlib.Path(__file__).parent / 'docs'


@pytest.fixture
def docs_dir():
    return docs


@pytest.fixture
def html_parser():
    return parsers['html']()


@pytest.fixture(params=sorted(docs.glob('*.html')), ids=lambda p: str(p.name))
def html_document(request):
    return request.param
