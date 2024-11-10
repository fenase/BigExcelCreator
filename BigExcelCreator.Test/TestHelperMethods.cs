using BigExcelCreator;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Test
{
    internal static class TestHelperMethods
    {
        internal static IEnumerable<Row> GetRows(Worksheet worksheet)
        {
            IEnumerable<SheetData> sheetDatas = worksheet.ChildElements.OfType<SheetData>();
            Assert.Multiple(() =>
            {
                Assert.That(sheetDatas, Is.Not.Null);
                Assert.That(sheetDatas.Count(), Is.EqualTo(1));
            });
            SheetData sheetData = sheetDatas.First();
            return sheetData.ChildElements.OfType<Row>();
        }

        internal static IEnumerable<Column> GetColumns(Worksheet worksheet)
        {
            IEnumerable<Columns> columnsData = worksheet.ChildElements.OfType<Columns>();
            Assert.Multiple(() =>
            {
                Assert.That(columnsData, Is.Not.Null);
                Assert.That(columnsData.Count(), Is.EqualTo(1));
            });
            Columns columns = columnsData.First();
            return columns.ChildElements.OfType<Column>();
        }

        internal static IEnumerable<ConditionalFormatting> GetConditionalFormatting(Worksheet worksheet)
        {
            IEnumerable<ConditionalFormatting> conditionalFormattingData = worksheet.ChildElements.OfType<ConditionalFormatting>();
            Assert.Multiple(() =>
            {
                Assert.That(conditionalFormattingData, Is.Not.Null);
            });
            return conditionalFormattingData;
        }

        internal static IEnumerable<Cell> GetCells(Row row)
        {
            return row.ChildElements.OfType<Cell>();
        }

        internal static string GetCellRealValue(Cell cell, WorkbookPart workbookPart)
        {
            switch (cell.DataType?.ToString())
            {
                case "s":
                    return workbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue!.Text.ToString()!)).Text!.Text;
                case "str":
                default:
                    return cell.CellValue!.Text;
            }
        }

        internal static BigExcelWriter GetWriterStream(out MemoryStream stream)
        {
            stream = new MemoryStream();
            return new BigExcelWriter(stream);
        }

        internal static IEnumerable<SpreadsheetDocumentType> ValidSpreadsheetDocumentTypes()
        {
            return
            [
                SpreadsheetDocumentType.Workbook,
                SpreadsheetDocumentType.Template,
                SpreadsheetDocumentType.MacroEnabledWorkbook,
                SpreadsheetDocumentType.MacroEnabledTemplate,
            ];
        }

        internal static IEnumerable<SpreadsheetDocumentType> InvalidSpreadsheetDocumentTypes()
        {
            return
            [
                SpreadsheetDocumentType.AddIn,
            ];
        }
    }
}
