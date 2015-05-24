import HtmlToWord
import win32com.client
import os
import sys
import pprint


word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = True
parser = HtmlToWord.Parser()


document = word.Documents.Add()

parser.ParseAndRender('<p>paragraph</p>', word, document.ActiveWindow.Selection)
parser.ParseAndRender('<h1>header</h1>', word, document.ActiveWindow.Selection)