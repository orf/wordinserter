WordInserter
===
This module allows you to insert HTML or MarkDown into a Word Document, as well as allowing you to programmatically build 
word documents in pure Python. The API is really simple to use:

``` python
from wordinserter import parse, render

operations = parse(html, parser="html") # or parser="markdown"
insert(operations, document=document, constants=constants)
```
    
Inserting HTML or Markdown into a Word document is a two step process: first the input has to be parsed into a sequence 
of operations, which is then *rendered* into a Word document. This library currently only supports inserting using the 
Word COM interface which means it is Windows specific at the moment.

Below is a more complex example including starting word that will insert a representation of the HTML code
into the new word document, including the image, caption and list.

``` python
from wordinserter import render, parse
from comtypes.client import CreateObject

# This opens Microsoft Word and creates a new document.
word = CreateObject("Word.Application")
word.Visible = True # Don't set this to True in production!
document = word.Documents.Add()
from comtypes.gen import Word as constants

html = """
<h3>This is a title</h3>
<p><img src="http://placehold.it/150x150" alt="I go below the image as a caption"></p>
<p><i>This is <b>some</b> text</i> in a <a href="http://google.com">paragraph</a></p>
<ul>
    <li>Boo! I am a <b>list</b></li>
</ul>
"""

# Parse the HTML into a list of operations then feed them into render.
operations = parse(html, parser="html")
render(operations, document=document, constants=constants)
```

What's with the constants part? Wordinserter is agnostic to the COM library you use. Each library exposes constant 
values that are needed by Wordinserter in a different way: the pywin32 library exposes it as win32com.client.constants 
whereas the comtypes library exposes them as a module that resides in comtypes.gen. Rather than guess which one you 
are using Wordinserter requires you to pass the right one in explicitly.


### Install
Get it [from PyPi here](https://pypi.python.org/pypi/wordinserter), using `pip install wordinserter`. This has been built with word 2010 and 2013, older 
versions may produce different results.


## Supported Operations
WordInserter currently supports a range of different operations, including code blocks, font size/colors, images, 
hyperlinks, numbered and bullet lists (
