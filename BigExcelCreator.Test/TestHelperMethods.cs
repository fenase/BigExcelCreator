using BigExcelCreator.Styles;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BigExcelCreator.Test
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

            Assert.That(conditionalFormattingData, Is.Not.Null);

            return conditionalFormattingData;
        }

        internal static IEnumerable<Cell> GetCells(Row row)
        {
            return row.ChildElements.OfType<Cell>();
        }

        internal static string GetCellRealValue(Cell cell, WorkbookPart workbookPart)
        {
            return (cell.DataType?.ToString()) switch
            {
                "s" => workbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue!.Text!)).Text!.Text,
                "str" or _ => cell.CellValue!.Text,
            };
        }

        internal static BigExcelWriter GetWriterStream(out MemoryStream stream)
        {
            stream = new MemoryStream();
            return new BigExcelWriter(stream);
        }



        internal static StyleList GetStyleList()
        {
            StyleList styleList = new();
            Font italic = new(new Italic());
            Font bold = new(new Bold());
            Font boldItalic = new(new Bold(), new Italic());
            styleList.NewStyle(italic, null, null, null, "italic default");
            styleList.NewStyle(bold, null, null, null, "bold default", out int boldStyleIndex);
            styleList.NewStyle(boldItalic, null, null, null, "bold italic default");

            Alignment center = new() { Horizontal = HorizontalAlignmentValues.Center };

            styleList.NewStyle(italic, null, null, null, center, "italic center");
            styleList.NewStyle(bold, null, null, null, center, "bold center");
            styleList.NewStyle(boldItalic, null, null, null, center, "bold italic center");
            Fill yellowFill = new Fill(new[]{
                        new PatternFill(new[]{
                            new ForegroundColor { Rgb = new HexBinaryValue { Value = "FFFF00" } } }
                        )
                        { PatternType = PatternValues.Solid } });
            styleList.NewStyle(null, yellowFill, null, null, "YELLOW");

            styleList.NewDifferentialStyle("RED", font: new Font(new[] { new Color { Rgb = new HexBinaryValue { Value = "FF0000" } } }));

            Fill greenFill = new Fill(new[]{
                                new PatternFill(new[]{
                                        new BackgroundColor { Rgb = new HexBinaryValue { Value = "00FF00" } } })
                        { PatternType = PatternValues.Solid } });

            styleList.NewDifferentialStyle("GREENBKG", fill: greenFill);
            return styleList;
        }
    }
}
