## 0.8.9 (WIP)
Don't do the formatting "x" hack unless there is actually a format to apply.

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