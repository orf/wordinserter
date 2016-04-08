WordInserter
===
This module allows you to insert HTML or MarkDown into a Word Document, as well as allowing you to programmatically build 
word documents in pure Python (Python 3.x only at the moment). The API is really simple to use:

``` python
from wordinserter import parse, insert

operations = parse(html, parser="html") # or parser="markdown"
insert(operations, document=document, constants=constants)
```
    
Inserting HTML or Markdown into a Word document is a two step process: first the input has to be parsed into a sequence 
of operations, which is then *inserted* into a Word document. This library currently only supports inserting using the 
Word COM interface which means it is Windows specific at the moment.

There is a [comparison document](https://rawgit.com/orf/wordinserter/master/Tests/report.html) showing the output of 
WordInserter against FireFox, check it out to see what the library can do.

Below is a more complex example including starting word that will insert a representation of the HTML code
into the new word document, including the image, caption and list.

``` python
from wordinserter import insert, parse
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

markdown = """
### This is a title

![I go below the image as a caption](http://placehold.it/150x150)

*This is **some** text* in a [paragraph](http://google.com)

  * Boo! I'm a **list**
"""

# Parse the HTML into a list of operations then feed them into insert.
# The Markdown can be parsed by using parser="markdown"
operations = parse(html, parser="html")
insert(operations, document=document, constants=constants)
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
hyperlinks, numbered and bullet lists.

#### Why aren't my lists showing up properly?
There are two ways people write lists in HTML, one with each sub-list as a child of the parent list, or as a child of a
list element. Below is a sample of the two different ways, both of which display correctly in all browsers:

```
<ol>
    <li>
        I'm a list element
    </li>
    <ul>
        <li>I'm a sub list!</li>
    </ul>
</ol>
```
```
<ol>
    <li>
        I'm a list element
        <ul>
            <li>I'm a sub list!</li>
        </ul>
    </li>
</ol>
```

The second way is correct according to the HTML specification. `lxml` parses the first structure incorrectly in some cases,
which leads to weird list behavior. There isn't much this library can do about that, so make sure your lists are
in the second format.