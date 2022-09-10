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





# Build and Test
TODO: Describe and show how to build your code and run the tests. 

# Contribute
TODO: Explain how other users and developers can contribute to make your code better. 

If you want to learn more about creating good readme files then refer the following [guidelines](https://docs.microsoft.com/en-us/azure/devops/repos/git/create-a-readme?view=azure-devops). You can also seek inspiration from the below readme files:
- [ASP.NET Core](https://github.com/aspnet/Home)
- [Visual Studio Code](https://github.com/Microsoft/vscode)
- [Chakra Core](https://github.com/Microsoft/ChakraCore)