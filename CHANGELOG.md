## 1.1.3

Add support for inserting page breaks

## 1.1.2
Fix hyperlinks/styles/code sections inside textboxes

Support 'margin-left' property for tables

Support table widths in points as well as percent

Allow STYLEREF field code

Add support for table borders

Support padding in tables and table cells

## 1.1.1
Add support for specifying the line-height attribute as a percentage

## 1.1.0
Add support for `line-height` CSS attribute

## 1.0.9
Bugfix for incredibly long hyperlinks

## 1.0.8
Bugfix for custom word fields like `'@DOCUMENT'`

## 1.0.7
Accept all mimetypes in data-uri images.

## 1.0.6
CSS `padding-top` and `padding-bottom` attributes are supported. This sets the `SpaceAfter` and `SpaceBefore` 
property on the `ParagraphFormat` object.

Hooks are now fired for `Format` operations as well, with extra arguments for the original operation and 
the range it encompasses.

## 1.0.5
Add support for more word fields in hyperlinks. The syntax is `<a href='@CODE'>`. Currently only `FILENAME` is supported

## 1.0.4
Add support for text-alignment in paragraphs.

## 1.0.3
Add support for table cell text orientations through the css `writing-mode` property

## 1.0.0
Add support for inserting cross references. Use `<a href='!ref'>`. If the reference does not exist yet you will need to
select the whole document and call Fields.Update() before saving.

Add a `wordinserter` CLI tool for quick prototyping

Correctly handle tables with mixed widths

## 0.9.6.1
Correctly handle 'pt' font-size CSS declarations. Since 0.9.5.1 they where incorrectly treated the same as 'px' values.

## 0.9.6
Set the default parser to be 'html' in `wordinserter.render`. Also improve handling of elements with invalid children,
so for example this markup correctly renders: `<h1><strong>Test</strong></h1>`

## 0.9.5.1
Support non-float CSS font-sizes

## 0.9.5
Correctly support the background color in display: block elements.

## 0.9.4.6
Actually fix headings not having bookmarks applied.

## 0.9.4.5
Fix headings not having bookmarks applied if they have an `id` attribute set.

## 0.9.4.4
Fix Heading element not being correctly rendered, and fix cell table widths throwing an error if they are 
pixel based but the parents are percentage based

## 0.9.4.3
Fix for some python 3.5 specific syntax

## 0.9.4.2
Ignore text elements from child validation

## 0.9.4.1
Don't error with images that have no src attribute.

## 0.9.4
Add support for rotated text in table cells

## 0.9.3
Correctly handle CSS inheritance: child styles are applied after the parent ones (before this they were
applied top down, children first). This means child styles correctly override parent ones.

## 0.9.2.7
Handle broken images correctly. If an image is invalid then the `404.png` will be inserted.

## 0.9.2.6
Handle hyperlinks after lists correctly.

## 0.9.2.5
Improve error support, InsertErrors now store the `exc_info()` of the 
inner exception.

## 0.9.2.4
Fix CodeBlock spacing styles being incorrectly reverted after the 
CodeBlock has finished.

## 0.9.2.3
Improve list formatting

## 0.9.2.2
Fix using int() when floats are expected.

## 0.9.2
Fix handling of background colors

## 0.9.1
Fix table cell range handling. 

## 0.9
Don't do the formatting "x" hack unless there is actually a format to apply. Improved the speed of table creation.
Disabled spell checking for code blocks. All nodes now support an "id" attribute, and if it is specified in a 
heading tag a bookmark is added with the ID attribute as its name. Hyperlinks also now support a bookmark if their 
URL's start with a "#", so `<a href="#name">` would link to `<h1 id="name">`.

Majorly refactored the whitespace handling code. It not actually works :)

## 0.8.8
Fix line spacing after CodeBlocks. Added support for border styles in images, and for multiple constants (use the
`CombinedConstants` class)

## 0.8.7
Added support for CSS files! WHERE IS YOUR GOD NOW? You can pass a `stylesheets` array to `parse()` with some css
definitions and it will do the right thing.

## 0.8.6
Added support for images with data URIs. Passing an image with a src set to "data:mimetype,base64;DATA..." will work
as expected, by extracting the image from the data URI and saving it, before inserting.

## 0.8.5
Remove `break_across_pages`. Word is horrible, you need to make the application Visible=True for
this to work. This should be handled by your app code not this library.

## 0.8.4
Add an attribute `break_across_pages` to the Table operation. Add this to make your tables break
across pages, hopefully.

## 0.8.3
Fix an encoding error on Windows Server. Why windows, why. :(

## 0.8.2
Fixed an error when an unknown language is given inside a code block.

## 0.8.1
Refactored the lists implementation to not suck and actually function. We now support roman numerals,
complex nested lists and other funky stuff. Added a `lxml` dependency and hard-coded the parser to be `lxml`.

## 0.8.0
Remove pywin32 dependency

## 0.7.9
Add support for table cells with background colors.

## 0.7.8
Actually handle ListElements with no children

## 0.7.7
Handle ListElements with no children

## 0.7.6
Apparently there is already a 0.7.5 release. Not sure how that happened.

## 0.7.5
Added support for CSS named colors, i.e `color: black`.

## 0.7.4
Callbacks are now called **after** the operation has been styled, instead of before.

## 0.7.3
Hopefully improved whitespace handling inside `ListElements`, and improved the logic behind `LineBreak` operations.

## 0.7.2
Fixes for table cells with no width, also always ensure a `Format` object is attached to an Operation.

## 0.7.1
Added support for tables with percentage widths.

## 0.6.13
Added support for text-align CSS classes inside tables

## 0.6.12
Added support for <br> tags.

## 0.6.11
Fixed an issue where inserting an image as the only content of a TableCell would cause it to be inserted in the wrong place.
