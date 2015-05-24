from HtmlToWord.parser.html import HTMLParser
from HtmlToWord.render.com import COMRenderer
import win32com.client

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = True
document = word.Documents.Add()
parser = HTMLParser()

HTML = """
<strong><i>Hello</i> <u>there</u></strong> Sonny jim

What can we do for ya?   <strong>Sup bro</strong>
"""

renderer = COMRenderer(word, document.ActiveWindow.Selection)
renderer.render(parser.parse(HTML))