# Introduction 
Create Excel files using OpenXML SAX.

Note: This package is a work in progress.

Styling is available, but instructions are pending.

# Usage
1. Instantiate class `BigExcelWritter` using either a file path or a stream (`MemoryStream` is recommended).
2. Open a new Sheet using `CreateAndOpenSheet`
3. For every row, use `BeginRow` and `EndRow`
4. Between `BeginRow` and `EndRow`, use `WriteTextCell` to write a cell.
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
    excel.CloseSheet();
}
```



# Styling and formatting
## Column width
When calling `CreateAndOpenSheet`, pass `IList<Column>` as second parameter.
Each element represents a single column.
Only the `CustomWidth` and `Width` properties are needed and considered.

> At this point, `CustomWidth` must be set to `true`

`Width` represents the column width in characters (Same unit as when resizing in Excel).

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
* `SheetStateValues.Visible` (default): Sheet is vissible
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
    excel.CreateAndOpenSheet("Sheet Name");
    excel.BeginRow();
    excel.WriteTextCell("This has a gray patterned background", name1);
    excel.WriteTextCell("This has a yellow backgound", name2);
    excel.EndRow();
    excel.CloseSheet();
}
```




# Build and Test
TODO: Describe and show how to build your code and run the tests. 

# Contribute
TODO: Explain how other users and developers can contribute to make your code better. 

If you want to learn more about creating good readme files then refer the following [guidelines](https://docs.microsoft.com/en-us/azure/devops/repos/git/create-a-readme?view=azure-devops). You can also seek inspiration from the below readme files:
- [ASP.NET Core](https://github.com/aspnet/Home)
- [Visual Studio Code](https://github.com/Microsoft/vscode)
- [Chakra Core](https://github.com/Microsoft/ChakraCore)