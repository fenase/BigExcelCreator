# Big Excel Creator

Create Excel files using OpenXML SAX with styling.
This is specially useful when trying to output thousands of rows.

The idea behind this package is to be a basic easy-to-use wrapper around
[DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml)
aimed towards generating large excel files as fast as possible using the SAX method for writing.

At the same time, this writer should prevent you from creating an invalid file
(i.e.: a file generated without any errors, but unable to be opened).
Since the most common reason for a file to become corrupted when creating it using SAX is out-of-order instructions
(i.e.: writing to a cell outside a sheet), this package should detect that, and throw an exception.

[![Nuget](https://img.shields.io/nuget/v/BigExcelCreator)](https://www.nuget.org/packages/BigExcelCreator)
[![Build Status](https://dev.azure.com/fenase/BigExcelCreator/_apis/build/status%2FBigExcelCreator-CI?branchName=main)](https://dev.azure.com/fenase/BigExcelCreator/_build/latest?definitionId=18&branchName=main)
[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=fenase_BigExcelCreator&metric=alert_status)](https://sonarcloud.io/summary/overall?id=fenase_BigExcelCreator)
[![Lines of Code](https://sonarcloud.io/api/project_badges/measure?project=fenase_BigExcelCreator&metric=ncloc)](https://sonarcloud.io/summary/overall?id=fenase_BigExcelCreator)
[![Coverage](https://sonarcloud.io/api/project_badges/measure?project=fenase_BigExcelCreator&metric=coverage)](https://sonarcloud.io/summary/overall?id=fenase_BigExcelCreator)


# Table of Contents

- [Usage](#usage)
  - [Shared Strings](#shared-strings)
- [Data Validation](#data-validation)
- [Styling and formatting](#styling-and-formatting)
  - [Column formatting](#column-formatting)
  - [Hide Sheet](#hide-sheet)
  - [Merge Cells](#merge-cells)
  - [Styling](#styling)
  - [Comments](#comments)
  - [Autofilter](#autofilter)
  - [Conditional Formatting](#conditional-formatting)
    - [Formula](#formula)
    - [Cell Is](#cell-is)
    - [Duplicated Values](#duplicated-values)
- [Page Layout](#page-layout)
  - [Sheet options](#sheet-options)
    - [Gridlines](#gridlines)
    - [Headings](#headings)


# Usage

1. Instantiate class `BigExcelWriter` using either a file path or a stream (`MemoryStream` is recommended).
2. Open a new Sheet using `CreateAndOpenSheet`
3. For every row, use `BeginRow` and `EndRow`
    * If you want to hide a row, pass `true` when calling `BeginRow`
4. Between `BeginRow` and `EndRow`, use `WriteTextCell` to write a cell.
    > Alternatively, you can use `WriteTextRow` to write an entire row at once, using the same format.
    
    > Starting on version 1.1, text cells can be written using the shared strings table, which should reduce the generated file size.
    > See [Shared Strings](#shared-strings) below
5. Use `WriteFormulaCell` or `WriteFormulaRow` to insert formulas.
6. Use `WriteNumberCell` or `WriteNumberRow` to insert numbers. This is useful if you need to do any calculation later on.
7. Use `CloseSheet` to finish.
8. If needed, repeat steps 2 -> 5 to write to another sheet

## Shared Strings

If the same text appears across different sheets, using the shared strings table may help reduce the generated file size.
In order to do this, simply set to `true` the `useSharedStrings` parameter when calling `WriteTextCell` or `WriteTextRow`.


## Example

```c#
using BigExcelCreator;

....

MemoryStream stream = new MemoryStream();
using (BigExcelWriter excel = new(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
{
    excel.CreateAndOpenSheet("Sheet Name");
    excel.BeginRow();
    excel.WriteTextCell("Cell content");
    excel.WriteTextCell(123); // write as number. This allows to use formulas.
    excel.WriteTextCell(456);
    excel.WriteFormulaCell("SUM(B1:C1)");
    excel.EndRow();
    excel.BeginRow(true);
    excel.WriteTextCell("This row is hidden");
    excel.EndRow();
    excel.CloseSheet();
}
```


# Data Validation

Use `AddListValidator` to restrict, to a list defined in a formula,
possible values to be written to a cell by an user.

Alternatively, use `AddIntegerValidator` or `AddDecimalValidator` to restrict / validate values
as defined by `validationType` (equal, greater than, between, etc.)

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


## Merge Cells

In order to merge a range of cells while a sheet is open, use `MergeCells` with a range.
```c#
excel.MergeCells("A1:A5");
```


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

When instantiating `BigExcelWriter`, use the result of calling `GetStylesheet` as the `stylesheet` parameter.
Then, when writing a cell, you can use the name given earlier to format it.

```c#
MemoryStream stream = new MemoryStream();
using (BigExcelWriter excel = new(stream,
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


> If you're planning to use [Conditional Formatting](#conditional-formatting), you must also create differential styles here.
> To do so, follow the same instructions as above, replacing `NewStyle` with `NewDifferentialStyle`.
> 
> All parameters of `NewDifferentialStyle` are optional, except `name`. Of the optional parameters, at least one must be present.

```c#
// place this before calling list.GetStylesheet() and new BigExcelWriter()
list.NewDifferentialStyle("RED", font: new Font(new Color { Rgb = new HexBinaryValue { Value = "FF0000" } }));
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

## Conditional Formatting

In order to use conditional formatting, you should define Differential styles (see [Styling](#styling))

> On every case below:
> - `reference` => A range of cells to apply the conditional formatting to
> - `format` => The id of the Differential style. Obtain it using `GetIndexDifferentialByName` after creating it with `NewDifferentialStyle`


### Formula

To define a conditional style by formula, use `AddConditionalFormattingFormula(string reference, string formula, int format)`.

- `formula` defines the expression to use.
    Use a fixed range using `$` to anchor the reference to a cell.
    Avoid using `$` to make the reference "walk" with the range.
    This is useful when referencing the current cell.

```c#
excel.AddConditionalFormattingFormula("A1:A10", "A1<5", styleList.GetIndexDifferentialByName("RED"));
```

### Cell Is

Format cells based on their contents using `AddConditionalFormattingCellIs`

- `Operator` defines how to compare values.
- `value` defines the value to compare the cell to.
- `value2` If the operator requires 2 numbers (eg: `Between` and `NotBetween`), the second value goes here.

```c#
excel.AddConditionalFormattingCellIs("A1:A20", ConditionalFormattingOperatorValues.LessThan, "5", styleList.GetIndexDifferentialByName("RED"));
excel.AddConditionalFormattingCellIs("A1:A20", ConditionalFormattingOperatorValues.Between, "3", styleList.GetIndexDifferentialByName("RED"), "7");
```

### Duplicated Values

Format duplicated values using `AddConditionalFormattingDuplicatedValues`

```c#
excel.AddConditionalFormattingDuplicatedValues("A1:A10", styleList.GetIndexDifferentialByName("RED"));
```


# Page Layout

## Sheet options

### Gridlines

While working on a sheet, the property `ShowGridLinesInCurrentSheet` controls whether the gridlines are shown.
Enabled by default.

While working on a sheet, the property `PrintGridLinesInCurrentSheet` controls whether the gridlines are printed.
Disabled by default.

### Headings

While working on a sheet, the property `ShowRowAndColumnHeadingsInCurrentSheet` controls whether the headings (Column letters and row numbers) are shown.
Enabled by default.

While working on a sheet, the property `PrintRowAndColumnHeadingsInCurrentSheet` controls whether the headings (Column letters and row numbers) are printed.
Disabled by default.
