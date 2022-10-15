using BigExcelCreator;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        internal static BigExcelwriter GetwriterStream(out MemoryStream stream)
        {
            stream = new MemoryStream();
            return new BigExcelwriter(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        }
    }
}
