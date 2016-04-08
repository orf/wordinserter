## 0.8.1 (WIP)
Remove hard-coded lxml parser in BeautifulSoup. Refactored the lists implementation to not suck and actually
function. We now support roman numerals, complex nested lists and other funky stuff.

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