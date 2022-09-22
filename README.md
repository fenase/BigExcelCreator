# Big Excel Creator

Create Excel files using OpenXML SAX with styling.
This is specially useful when trying to output thousands of rows

[![Nuget](https://img.shields.io/nuget/v/BigExcelCreator)](https://www.nuget.org/packages/BigExcelCreator)
[![Build status](https://dev.azure.com/fenase/BigExcelCreator/_apis/build/status/BigExcelCreator-CI)](https://dev.azure.com/fenase/BigExcelCreator/_build/latest?definitionId=4)

# Table of Contents

- [Usage](#usage)
- [Data Validation](#data-validation)
- [Styling and formatting](#styling-and-formatting)
    - [Column formatting](#column-formatting)
    - [Hide Sheet](#hide-sheet)
    - [Styling](#styling)
    - [Comments](#comments)
    - [Autofilter](#autofilter)


# Usage

1. Instantiate class `BigExcelWritter` using either a file path or a stream (`MemoryStream` is recommended).
2. Open a new Sheet using `CreateAndOpenSheet`
3. For every row, use `BeginRow` and `EndRow`
    * If you want to hide a row, pass `true` when calling `BeginRow`
4. Between `BeginRow` and `EndRow`, use `WriteTextCell` to write a cell.
    > Alternatively, you can use `WriteTextRow` to write an entire row at once, using the same format.
5. Use `CloseSheet` to finish.
6. If needed, repeat steps 2 -> 5 to write to another sheet

## Example

```c#
using BigExcelCreator;

....

MemoryStream stream = new MemoryStream();
using (BigExcelWritter excel = new(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
{
    excel.CreateAndOpenSheet("Sheet Name");
    excel.BeginRow();
    excel.WriteTextCell("Cell content");
    excel.EndRow();
    excel.BeginRow(true);
    excel.WriteTextCell("This row id hidden");
    excel.EndRow();
    excel.CloseSheet();
}
```


# Data Validation

Use `AddListValidator` to restrict possible values to be written to a cell by an user.
```c#
    excel.CreateAndOpenSheet("Sheet Name");
    
    ...    
    
    // Only allow values included in sheet named "vals" between cells A1 and A6
    // when writing to cells between B2 and B10 of the current sheet.
    string range = "B2:B10";
    string formula = "vals!$A$1:$A$6";
    excel.AddValidator(range, formula);
    
    excel.CloseSheet();
```



# Styling and formatting

## Column formatting

When calling `CreateAndOpenSheet`, pass `IList<Column>` as second parameter.
Each element represents a single column.
Only the `CustomWidth`, `Width` and `Hidden` are used.

`Width` represents the column width in characters (Same unit as when resizing in Excel).

`CustomWidth` allows the use of `Width`.

`Hidden` hides the column.

### Example

```c#
List<Column> cols = new List<Column> {
    new Column{CustomWidth = true, Width=10},   // A
    new Column{CustomWidth = true, Width=15},   // B
    new Column{CustomWidth = true, Width=18},   // c
};

excel.CreateAndOpenSheet("Sheet Name", cols);

```


## Hide Sheet

`CreateAndOpenSheet` accepts as third parameter a `SheetStateValues` variable.
* `SheetStateValues.Visible` (default): Sheet is visible
* `SheetStateValues.Hidden`: Sheet is hidden
* `SheetStateValues.VeryHidden`: Sheet is hidden and cannot be unhidden from Excel's UI.


## Styling

First, the elements that define a style (font, fill, border and, optionally, numbering format) must be created.
```c#
font1 = new Font(new Bold(),
            new FontSize { Val = 11 },
            new Color { Rgb = new HexBinaryValue { Value = "000000" } },
            new FontName { Val = "Calibri" });

fill1 = new Fill(
            new PatternFill { PatternType = PatternValues.Gray125 });
fill2 = new Fill(
            new PatternFill (
                new ForegroundColor { Rgb = new HexBinaryValue { Value = "FFFF00" } }
            )
            { PatternType = PatternValues.Solid });

border1 = new Border(
            new LeftBorder(
                new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } }
            )
            { Style = BorderStyleValues.Thin },
            new RightBorder(
                new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } }
            )
            { Style = BorderStyleValues.Thin },
            new TopBorder(
                new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } }
            )
            { Style = BorderStyleValues.Thin },
            new BottomBorder(
                new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } }
            )
            { Style = BorderStyleValues.Thin },
            new DiagonalBorder());

numberingFormat1 = new NumberingFormat { NumberFormatId = 164, FormatCode = "0,.00;(0,.00)" };
```

After that, a new style list can be created and new styles inserted. Remember to name you styles.
```c#
StyleList list = new StyleList();
string name1 = "name1";
string name2 = "name2";

list.NewStyle(font1, fill1, border1, numberingFormat1, name1);
list.NewStyle(font1, fill2, border1, numberingFormat1, name2);
```

When instantiating `BigExcelWritter`, use the result of calling `GetStylesheet` as the `stylesheet` parameter.
Then, when writing a cell, you can use the name given earlier to format it.

```c#
MemoryStream stream = new MemoryStream();
using (BigExcelWritter excel = new(stream,
                                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                                    stylesheet: list.GetStylesheet()))
{
    int index_style_name1 = list.GetIndexByName(name1);
    int index_style_name2 = list.GetIndexByName(name2);
    excel.CreateAndOpenSheet("Sheet Name");
    excel.BeginRow();
    excel.WriteTextCell("This has a gray patterned background", index_style_name1);
    excel.WriteTextCell("This has a yellow background", index_style_name2);
    excel.EndRow();
    excel.CloseSheet();
}
```


## Comments

In order to add a note (formerly known as comment) to a cell, while a sheet is open, call the `Comment` method.

```c#
excel.CreateAndOpenSheet("Sheet Name");
excel.BeginRow();

excel.WriteTextCell("This has a gray patterned background", index_style_name1);
excel.WriteTextCell("This has a yellow background", index_style_name2);

excel.Comment("test A1 another sheet", "A1");

excel.EndRow();

excel.Comment("test E2 another sheet", "B1", "Author");

excel.CloseSheet();
```

## Autofilter

In order to add an Autofilter, call `AddAutofilter` while on a sheet.
```c#
excel.BeginRow();
// ...
excel.AddAutofilter(range); // Range's height must be 1. Example: A1:J1
// ...
excel.EndRow();
```

