---
_layout: landing
---

# BigExcelCreator

Create Excel files using OpenXML SAX with styling.
This is specially useful when trying to output thousands of rows.

[Package site](https://fenase.github.io/projects/BigExcelCreator)

[Get on NuGet](https://www.nuget.org/packages/BigExcelCreator/)

[API](/BigExcelCreator/api/BigExcelCreator.html)


The idea behind this package is to be a basic easy-to-use wrapper around a reduced subset of functions of 
[DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml)
aimed towards generating large excel files as fast as possible using the SAX method for writing.

At the same time, this writer should prevent you from creating an invalid file
(i.e.: a file generated without any errors, but unable to be opened).
Since the most common reason for a file to become corrupted when creating it using SAX is out-of-order instructions
(i.e.: writing to a cell outside a sheet), this package should detect that, and throw an exception.


