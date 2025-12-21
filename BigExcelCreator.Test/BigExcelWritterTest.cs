// Ignore Spelling: Validator

using BigExcelCreator.Exceptions;
using BigExcelCreator.Ranges;
using BigExcelCreator.Styles;
using BigExcelCreator.Test.Classes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using static BigExcelCreator.Test.TestHelperMethods;

namespace BigExcelCreator.Test
{
    internal class BigExcelWriterTest
    {
        private static readonly string DirectoryPath = TestContext.CurrentContext.WorkDirectory + @"\excelOut";

        private static readonly IEnumerable<SpreadsheetDocumentType> ValidSpreadsheetDocumentTypes =
        [
            SpreadsheetDocumentType.Workbook,
            SpreadsheetDocumentType.Template,
            SpreadsheetDocumentType.MacroEnabledWorkbook,
            SpreadsheetDocumentType.MacroEnabledTemplate,
        ];

        private static readonly IEnumerable<SpreadsheetDocumentType> InvalidSpreadsheetDocumentTypes =
        [
            SpreadsheetDocumentType.AddIn,
        ];

        [SetUp]
        public void Setup()
        {
            Directory.CreateDirectory(DirectoryPath);
            Assert.That(DirectoryPath, Does.Exist);
        }

        [TearDown]
        public void TearDown()
        {
            new DirectoryInfo(DirectoryPath).Delete(true);
        }

        [TestCaseSource(nameof(ValidSpreadsheetDocumentTypes))]
        public void FileExistsAfterCreation(SpreadsheetDocumentType spreadsheetDocumentType)
        {
            string path = Path.Combine(DirectoryPath, $"{Guid.NewGuid()}.xlsx");
            using (BigExcelWriter writer = new(path, spreadsheetDocumentType))
            {
                // do nothing
            }
            Assert.That(path, Does.Exist);
        }

        [Test]
        public void FileExistsAfterCreationDefaultDocumentType()
        {
            string path = Path.Combine(DirectoryPath, $"{Guid.NewGuid()}.xlsx");
            using (BigExcelWriter writer = new(path))
            {
                // do nothing
            }
            Assert.That(path, Does.Exist);
        }

        [TestCaseSource(nameof(InvalidSpreadsheetDocumentTypes))]
        public void UnsupportedSpreadsheetDocumentTypeFileCreationFails(SpreadsheetDocumentType spreadsheetDocumentType)
        {
            Assert.Multiple(() =>
            {
                string path = Path.Combine(DirectoryPath, $"{Guid.NewGuid()}.xlsx");

                Assert.Throws<UnsupportedSpreadsheetDocumentTypeException>(() =>
                {
                    using BigExcelWriter writer = new(path, spreadsheetDocumentType);
                    // do nothing
                });
                Assert.That(path, Does.Not.Exist);
            });

            Assert.Multiple(() =>
            {
                string path = Path.Combine(DirectoryPath, $"{Guid.NewGuid()}.xlsx");

                Assert.Catch<NotSupportedException>(() =>
                {
                    using BigExcelWriter writer = new(path, spreadsheetDocumentType);
                    // do nothing
                });
                Assert.That(path, Does.Not.Exist);
            });
        }

        [TestCaseSource(nameof(InvalidSpreadsheetDocumentTypes))]
        public void UnsupportedSpreadsheetDocumentTypeStreamCreationFails(SpreadsheetDocumentType spreadsheetDocumentType)
        {
            Assert.Multiple(() =>
            {
                Stream stream = new MemoryStream();
                Assert.Throws<UnsupportedSpreadsheetDocumentTypeException>(() =>
                {
                    using BigExcelWriter writer = new(stream, spreadsheetDocumentType);
                    // do nothing
                });
            });

            Assert.Multiple(() =>
            {
                Stream stream = new MemoryStream();
                Assert.Catch<NotSupportedException>(() =>
                {
                    using BigExcelWriter writer = new(stream, spreadsheetDocumentType);
                    // do nothing
                });
            });
        }

        [TestCaseSource(nameof(ValidSpreadsheetDocumentTypes))]
        public void ValidFile(SpreadsheetDocumentType spreadsheetDocumentType)
        {
            string path = Path.Combine(DirectoryPath, $"{Guid.NewGuid()}.xlsx");
            using (BigExcelWriter writer = new(path, spreadsheetDocumentType))
            {
                writer.CreateAndOpenSheet("first");
                writer.CloseSheet();
            }
            Assert.That(path, Does.Exist);

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(path, false);

            Assert.That(reader.DocumentType, Is.EqualTo(spreadsheetDocumentType));

            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Workbook workbook = workbookPart!.Workbook;

            Sheets? sheets = workbook.Sheets;
            Assert.Multiple(() =>
            {
                Assert.That(sheets, Is.Not.Null);
                Assert.That(sheets!.Count(), Is.EqualTo(1));
            });
            Sheet sheet = (Sheet)sheets!.ChildElements[0];
            Assert.Multiple(() =>
            {
                Assert.That(sheet, Is.Not.Null);
                Assert.That(sheet.Name!.ToString(), Is.EqualTo("first"));
            });
        }

        [Test]
        public void ValidFileDefaultDocumentType()
        {
            string path = Path.Combine(DirectoryPath, $"{Guid.NewGuid()}.xlsx");
            using (BigExcelWriter writer = new(path))
            {
                writer.CreateAndOpenSheet("first");
                writer.CloseSheet();
            }
            Assert.That(path, Does.Exist);

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(path, false);

            Assert.That(reader.DocumentType, Is.EqualTo(SpreadsheetDocumentType.Workbook));

            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Workbook workbook = workbookPart!.Workbook;

            Sheets? sheets = workbook.Sheets;
            Assert.Multiple(() =>
            {
                Assert.That(sheets, Is.Not.Null);
                Assert.That(sheets!.Count(), Is.EqualTo(1));
            });
            Sheet sheet = (Sheet)sheets!.ChildElements[0];
            Assert.Multiple(() =>
            {
                Assert.That(sheet, Is.Not.Null);
                Assert.That(sheet.Name!.ToString(), Is.EqualTo("first"));
            });
        }

        [TestCaseSource(nameof(ValidSpreadsheetDocumentTypes))]
        public void ValidStream(SpreadsheetDocumentType spreadsheetDocumentType)
        {
            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream, spreadsheetDocumentType))
            {
                writer.CreateAndOpenSheet("first");
                writer.CloseSheet();
            }
            Assert.Multiple(() =>
            {
                Assert.That(stream.Position, Is.Zero);
                Assert.That(stream, Has.Length.GreaterThan(0));
            });

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false);

            Assert.That(reader.DocumentType, Is.EqualTo(spreadsheetDocumentType));

            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Workbook workbook = workbookPart!.Workbook;

            Sheets? sheets = workbook.Sheets;
            Assert.Multiple(() =>
            {
                Assert.That(sheets, Is.Not.Null);
                Assert.That(sheets!.Count(), Is.EqualTo(1));
            });
            Sheet sheet = (Sheet)sheets!.ChildElements[0];
            Assert.Multiple(() =>
            {
                Assert.That(sheet, Is.Not.Null);
                Assert.That(sheet.Name!.ToString(), Is.EqualTo("first"));
            });
        }

        [Test]
        public void ValidStreamDefaultDocumentType()
        {
            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream))
            {
                writer.CreateAndOpenSheet("first");
                writer.CloseSheet();
            }
            Assert.Multiple(() =>
            {
                Assert.That(stream.Position, Is.Zero);
                Assert.That(stream, Has.Length.GreaterThan(0));
            });

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false);

            Assert.That(reader.DocumentType, Is.EqualTo(SpreadsheetDocumentType.Workbook));

            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Workbook workbook = workbookPart!.Workbook;

            Sheets? sheets = workbook.Sheets;
            Assert.Multiple(() =>
            {
                Assert.That(sheets, Is.Not.Null);
                Assert.That(sheets!.Count(), Is.EqualTo(1));
            });
            Sheet sheet = (Sheet)sheets!.ChildElements[0];
            Assert.Multiple(() =>
            {
                Assert.That(sheet, Is.Not.Null);
                Assert.That(sheet.Name!.ToString(), Is.EqualTo("first"));
            });
        }

        [Test]
        public void ValidDefaultStyles()
        {
            MemoryStream stream;
            using (BigExcelWriter writer = GetWriterStream(out stream))
            {
                writer.CreateAndOpenSheet("first");
                writer.CloseSheet();
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false);

            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            WorkbookStylesPart? workbookStylesPart = workbookPart!.WorkbookStylesPart;
            Assert.That(workbookStylesPart, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(workbookStylesPart!.Stylesheet, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.ChildElements, Is.Empty);
                Assert.That(workbookStylesPart!.Stylesheet.Fonts, Is.Null);
                Assert.That(workbookStylesPart!.Stylesheet.Fills, Is.Null);
                Assert.That(workbookStylesPart!.Stylesheet.Borders, Is.Null);
                Assert.That(workbookStylesPart!.Stylesheet.CellFormats, Is.Null);
                Assert.That(workbookStylesPart!.Stylesheet.DifferentialFormats, Is.Null);
            });
        }

        [Test]
        public void ValidCustomUnmodifiedStyles()
        {
            StyleList styleList = new();

            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream, styleList.GetStylesheet()))
            {
                writer.CreateAndOpenSheet("first");
                writer.CloseSheet();
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false);

            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            WorkbookStylesPart? workbookStylesPart = workbookPart!.WorkbookStylesPart;
            Assert.That(workbookStylesPart, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(workbookStylesPart!.Stylesheet, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.ChildElements, Is.Not.Empty);
                Assert.That(workbookStylesPart!.Stylesheet.Fonts, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.Fills, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.Borders, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.CellFormats, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.CellFormats!.ChildElements, Has.Count.EqualTo(2)); // default formats (# = 2)
                Assert.That(workbookStylesPart!.Stylesheet.DifferentialFormats, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.DifferentialFormats!.ChildElements, Is.Empty);
            });
        }

        [Test]
        public void ValidCustomStyles()
        {
            StyleList styleList = new();
            Font italic = new(new Italic());
            Font bold = new(new Bold());
            Font boldItalic = new(new Bold(), new Italic());
            styleList.NewStyle(italic, null, null, null, "italic default");
            styleList.NewStyle(bold, null, null, null, "bold default");
            styleList.NewStyle(boldItalic, null, null, null, "bold italic default");

            Alignment center = new() { Horizontal = HorizontalAlignmentValues.Center };

            styleList.NewStyle(italic, null, null, null, center, "italic center");
            styleList.NewStyle(bold, null, null, null, center, "bold center");
            styleList.NewStyle(boldItalic, null, null, null, center, "bold italic center");
            Fill yellowFill = new(new[]{
                        new PatternFill(new[]{
                            new ForegroundColor { Rgb = new HexBinaryValue { Value = "FFFF00" } } }
                        )
                        { PatternType = PatternValues.Solid } });
            styleList.NewStyle(null, yellowFill, null, null, "YELLOW");

            styleList.NewDifferentialStyle("RED", font: new Font(new[] { new Color { Rgb = new HexBinaryValue { Value = "FF0000" } } }));

            Fill greenFill = new(new[]{
                        new PatternFill(new[]{
                            new BackgroundColor { Rgb = new HexBinaryValue { Value = "00FF00" } } })
                        { PatternType = PatternValues.Solid } });

            styleList.NewDifferentialStyle("GREENBKG", fill: greenFill);

            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream, styleList.GetStylesheet()))
            {
                writer.CreateAndOpenSheet("first");
                writer.CloseSheet();
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false);

            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            WorkbookStylesPart? workbookStylesPart = workbookPart!.WorkbookStylesPart;
            Assert.That(workbookStylesPart, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(workbookStylesPart!.Stylesheet, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.ChildElements, Is.Not.Empty);
                Assert.That(workbookStylesPart!.Stylesheet.Fonts, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.Fills, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.Borders, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.CellFormats, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.CellFormats!.ChildElements, Has.Count.EqualTo(9)); // default formats (# = 2) + inserted (# = 7)
                Assert.That(workbookStylesPart!.Stylesheet.DifferentialFormats, Is.Not.Null);
                Assert.That(workbookStylesPart!.Stylesheet.DifferentialFormats!.ChildElements, Has.Count.EqualTo(2));
            });
        }

        [Test]
        public void LargeFile()
        {
            Assert.That(() =>
            {
                using BigExcelWriter writer = GetWriterStream(out _);
                Random rng = new();

                writer.CreateAndOpenSheet("a");
                for (int i = 0; i < 10000; i++)
                {
                    writer.BeginRow();
                    for (int j = 0; j < 10; j++)
                    {
                        writer.WriteTextCell(rng.Next(0, 100).ToString(CultureInfo.InvariantCulture), useSharedStrings: true);
                    }
                    writer.EndRow();
                }
                writer.CloseSheet();
                writer.CreateAndOpenSheet("b");
                for (int i = 0; i < 10000; i++)
                {
                    writer.BeginRow();
                    for (int j = 0; j < 10; j++)
                    {
                        writer.WriteTextCell(rng.Next(0, 100).ToString(CultureInfo.InvariantCulture), useSharedStrings: true);
                    }
                    writer.EndRow();
                }
                writer.CloseSheet();
            }
            , Throws.Nothing);
        }

        [TestCaseSource(nameof(ValidSpreadsheetDocumentTypes))]
        public void ValidContent(SpreadsheetDocumentType spreadsheetDocumentType)
        {
            List<Column> creationColumns = [new Column { Width = 15 }, new Column { Width = 20 },];
            StyleList styleList = new();
            styleList.NewDifferentialStyle("RED", font: new Font(new[] { new Color { Rgb = new HexBinaryValue { Value = "FF0000" } } }));
            styleList.NewDifferentialStyle("GREEN", font: new Font(new[] { new Color { Rgb = new HexBinaryValue { Value = "00FF00" } } }));

            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream, spreadsheetDocumentType))
            {
                writer.CreateAndOpenSheet("first", creationColumns);
                writer.WriteTextRow(["a", "b", "c"]);
                writer.WriteNumberRow(new List<float> { 1f, 2f, 30f, 40f });
                writer.WriteFormulaRow(["SUM(A2:D2)"]);

                writer.AddConditionalFormattingCellIs("B1:B4", ConditionalFormattingOperatorValues.LessThan, "5", styleList.GetIndexDifferentialByName("RED"));
                writer.AddConditionalFormattingCellIs("B1:B4", ConditionalFormattingOperatorValues.Between, "3", styleList.GetIndexDifferentialByName("GREEN"), "7");

                writer.CloseSheet();
            }

            Assert.Multiple(() =>
            {
                Assert.That(stream.Position, Is.Zero);
                Assert.That(stream, Has.Length.GreaterThan(0));
            });

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false);

            Assert.That(reader.DocumentType, Is.EqualTo(spreadsheetDocumentType));

            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Workbook workbook = workbookPart!.Workbook;

            Sheets? sheets = workbook.Sheets;
            Assert.Multiple(() =>
            {
                Assert.That(sheets, Is.Not.Null);
                Assert.That(sheets!.Count(), Is.EqualTo(1));
            });
            Sheet sheet = (Sheet)sheets!.ChildElements[0];
            Assert.Multiple(() =>
            {
                Assert.That(sheet, Is.Not.Null);
                Assert.That(sheet.Name!.ToString(), Is.EqualTo("first"));
            });

            IEnumerable<Column> columns = GetColumns(workbookPart.WorksheetParts.First().Worksheet);
            Assert.Multiple(() =>
            {
                Assert.That(columns, Is.Not.Null);
                Assert.That(columns!.Count(), Is.EqualTo(creationColumns.Count));
                for (int i = 0; i < creationColumns.Count; i++)
                {
                    Assert.That(columns.ElementAt(i).CustomWidth, Is.Not.Null);
                    Assert.That(columns.ElementAt(i).CustomWidth?.Value, Is.True);
                    Assert.That(columns.ElementAt(i).Width, Is.EqualTo(creationColumns[i].Width));
                }
            });

            IEnumerable<ConditionalFormatting> conditionalFormattings = GetConditionalFormatting(workbookPart.WorksheetParts.First().Worksheet);
            Assert.Multiple(() =>
            {
                Assert.That(conditionalFormattings, Has.Exactly(2).Matches<ConditionalFormatting>(c => string.Equals(c.SequenceOfReferences, "B1:B4", StringComparison.InvariantCultureIgnoreCase)));
                Assert.That(conditionalFormattings, Has.Exactly(1).Matches<ConditionalFormatting>
                    (c => c.ChildElements.First<ConditionalFormattingRule>()!.Operator! == ConditionalFormattingOperatorValues.LessThan
                        && c.ChildElements.First<ConditionalFormattingRule>()!.FirstChild!.InnerText == "5"
                        && c.ChildElements.First<ConditionalFormattingRule>()!.LastChild!.InnerText == "5"));
                Assert.That(conditionalFormattings, Has.Exactly(1).Matches<ConditionalFormatting>
                    (c => c.ChildElements.First<ConditionalFormattingRule>()!.Operator! == ConditionalFormattingOperatorValues.Between
                        && c.ChildElements.First<ConditionalFormattingRule>()!.FirstChild!.InnerText == "3"
                        && c.ChildElements.First<ConditionalFormattingRule>()!.LastChild!.InnerText == "7"));
            });

            IEnumerable<Row> rows = GetRows(workbookPart.WorksheetParts.First().Worksheet);
            Assert.Multiple(() =>
            {
                Assert.That(rows, Is.Not.Null);
                Assert.That(rows.Count(), Is.EqualTo(3));
            });

            int skipRows = 0;
            IEnumerable<Cell> cells = GetCells(rows.Skip(skipRows).First());
            Assert.Multiple(() =>
            {
                Assert.That(cells, Is.Not.Null);
                Assert.That(cells.Count(), Is.EqualTo(3));
                Assert.That(GetCellRealValue(cells.Skip(0).Take(1).First(), workbookPart), Is.EqualTo("a"));
                Assert.That(GetCellRealValue(cells.Skip(1).Take(1).First(), workbookPart), Is.EqualTo("b"));
                Assert.That(GetCellRealValue(cells.Skip(2).Take(1).First(), workbookPart), Is.EqualTo("c"));
            });
            skipRows++;

            cells = GetCells(rows.Skip(skipRows).First());
            Assert.Multiple(() =>
            {
                Assert.That(cells, Is.Not.Null);
                Assert.That(cells.Count(), Is.EqualTo(4));
                Assert.That(GetCellRealValue(cells.Skip(0).Take(1).First(), workbookPart), Is.EqualTo("1"));
                Assert.That(GetCellRealValue(cells.Skip(1).Take(1).First(), workbookPart), Is.EqualTo("2"));
                Assert.That(GetCellRealValue(cells.Skip(2).Take(1).First(), workbookPart), Is.EqualTo("30"));
                Assert.That(GetCellRealValue(cells.Skip(3).Take(1).First(), workbookPart), Is.EqualTo("40"));
            });
            skipRows++;

            cells = GetCells(rows.Skip(skipRows).First());
            Assert.Multiple(() =>
            {
                Assert.That(cells, Is.Not.Null);
                Assert.That(cells.Count(), Is.EqualTo(1));
                Assert.That(cells.Skip(0).Take(1).First().CellFormula!.Text, Is.EqualTo("SUM(A2:D2)"));
            });
        }

        [Test]
        public void ValidContentDefaultDocumentType()
        {
            List<Column> creationColumns = [new Column { Width = 15 }, new Column { Width = 20 },];
            StyleList styleList = new();
            styleList.NewDifferentialStyle("RED", font: new Font(new[] { new Color { Rgb = new HexBinaryValue { Value = "FF0000" } } }));
            styleList.NewDifferentialStyle("GREEN", font: new Font(new[] { new Color { Rgb = new HexBinaryValue { Value = "00FF00" } } }));

            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream))
            {
                writer.CreateAndOpenSheet("first", creationColumns);
                writer.WriteTextRow(["a", "b", "c"]);
                writer.WriteNumberRow(new List<float> { 1f, 2f, 30f, 40f });
                writer.WriteFormulaRow(["SUM(A2:D2)"]);

                writer.AddConditionalFormattingCellIs("B1:B4", ConditionalFormattingOperatorValues.LessThan, "5", styleList.GetIndexDifferentialByName("RED"));
                writer.AddConditionalFormattingCellIs("B1:B4", ConditionalFormattingOperatorValues.Between, "3", styleList.GetIndexDifferentialByName("GREEN"), "7");

                writer.CloseSheet();
            }

            Assert.Multiple(() =>
            {
                Assert.That(stream.Position, Is.Zero);
                Assert.That(stream, Has.Length.GreaterThan(0));
            });

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false);

            Assert.That(reader.DocumentType, Is.EqualTo(SpreadsheetDocumentType.Workbook));

            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Workbook workbook = workbookPart!.Workbook;

            Sheets? sheets = workbook.Sheets;
            Assert.Multiple(() =>
            {
                Assert.That(sheets, Is.Not.Null);
                Assert.That(sheets!.Count(), Is.EqualTo(1));
            });
            Sheet sheet = (Sheet)sheets!.ChildElements[0];
            Assert.Multiple(() =>
            {
                Assert.That(sheet, Is.Not.Null);
                Assert.That(sheet.Name!.ToString(), Is.EqualTo("first"));
            });

            IEnumerable<Column> columns = GetColumns(workbookPart.WorksheetParts.First().Worksheet);
            Assert.Multiple(() =>
            {
                Assert.That(columns, Is.Not.Null);
                Assert.That(columns!.Count(), Is.EqualTo(creationColumns.Count));
                for (int i = 0; i < creationColumns.Count; i++)
                {
                    Assert.That(columns.ElementAt(i).CustomWidth, Is.Not.Null);
                    Assert.That(columns.ElementAt(i).CustomWidth?.Value, Is.True);
                    Assert.That(columns.ElementAt(i).Width, Is.EqualTo(creationColumns[i].Width));
                }
            });

            IEnumerable<ConditionalFormatting> conditionalFormattings = GetConditionalFormatting(workbookPart.WorksheetParts.First().Worksheet);
            Assert.Multiple(() =>
            {
                Assert.That(conditionalFormattings, Has.Exactly(2).Matches<ConditionalFormatting>(c => string.Equals(c.SequenceOfReferences, "B1:B4", StringComparison.InvariantCultureIgnoreCase)));
                Assert.That(conditionalFormattings, Has.Exactly(1).Matches<ConditionalFormatting>
                    (c => c.ChildElements.First<ConditionalFormattingRule>()!.Operator! == ConditionalFormattingOperatorValues.LessThan
                        && c.ChildElements.First<ConditionalFormattingRule>()!.FirstChild!.InnerText == "5"
                        && c.ChildElements.First<ConditionalFormattingRule>()!.LastChild!.InnerText == "5"));
                Assert.That(conditionalFormattings, Has.Exactly(1).Matches<ConditionalFormatting>
                    (c => c.ChildElements.First<ConditionalFormattingRule>()!.Operator! == ConditionalFormattingOperatorValues.Between
                        && c.ChildElements.First<ConditionalFormattingRule>()!.FirstChild!.InnerText == "3"
                        && c.ChildElements.First<ConditionalFormattingRule>()!.LastChild!.InnerText == "7"));
            });

            IEnumerable<Row> rows = GetRows(workbookPart.WorksheetParts.First().Worksheet);
            Assert.Multiple(() =>
            {
                Assert.That(rows, Is.Not.Null);
                Assert.That(rows.Count(), Is.EqualTo(3));
            });

            int skipRows = 0;
            IEnumerable<Cell> cells = GetCells(rows.Skip(skipRows).First());
            Assert.Multiple(() =>
            {
                Assert.That(cells, Is.Not.Null);
                Assert.That(cells.Count(), Is.EqualTo(3));
                Assert.That(GetCellRealValue(cells.Skip(0).Take(1).First(), workbookPart), Is.EqualTo("a"));
                Assert.That(GetCellRealValue(cells.Skip(1).Take(1).First(), workbookPart), Is.EqualTo("b"));
                Assert.That(GetCellRealValue(cells.Skip(2).Take(1).First(), workbookPart), Is.EqualTo("c"));
            });
            skipRows++;

            cells = GetCells(rows.Skip(skipRows).First());
            Assert.Multiple(() =>
            {
                Assert.That(cells, Is.Not.Null);
                Assert.That(cells.Count(), Is.EqualTo(4));
                Assert.That(GetCellRealValue(cells.Skip(0).Take(1).First(), workbookPart), Is.EqualTo("1"));
                Assert.That(GetCellRealValue(cells.Skip(1).Take(1).First(), workbookPart), Is.EqualTo("2"));
                Assert.That(GetCellRealValue(cells.Skip(2).Take(1).First(), workbookPart), Is.EqualTo("30"));
                Assert.That(GetCellRealValue(cells.Skip(3).Take(1).First(), workbookPart), Is.EqualTo("40"));
            });
            skipRows++;

            cells = GetCells(rows.Skip(skipRows).First());
            Assert.Multiple(() =>
            {
                Assert.That(cells, Is.Not.Null);
                Assert.That(cells.Count(), Is.EqualTo(1));
                Assert.That(cells.Skip(0).Take(1).First().CellFormula!.Text, Is.EqualTo("SUM(A2:D2)"));
            });
        }

        [Test]
        public void ValidContentDefaultColumns()
        {
            StyleList styleList = new();
            styleList.NewDifferentialStyle("RED", font: new Font(new[] { new Color { Rgb = new HexBinaryValue { Value = "FF0000" } } }));
            styleList.NewDifferentialStyle("GREEN", font: new Font(new[] { new Color { Rgb = new HexBinaryValue { Value = "00FF00" } } }));

            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream))
            {
                writer.CreateAndOpenSheet("first");
                writer.WriteTextRow(["a", "b", "c"]);
                writer.WriteNumberRow(new List<float> { 1f, 2f, 30f, 40f });
                writer.WriteFormulaRow(["SUM(A2:D2)"]);

                writer.AddConditionalFormattingCellIs("B1:B4", ConditionalFormattingOperatorValues.LessThan, "5", styleList.GetIndexDifferentialByName("RED"));
                writer.AddConditionalFormattingCellIs("B1:B4", ConditionalFormattingOperatorValues.Between, "3", styleList.GetIndexDifferentialByName("GREEN"), "7");

                writer.CloseSheet();
            }

            Assert.Multiple(() =>
            {
                Assert.That(stream.Position, Is.Zero);
                Assert.That(stream, Has.Length.GreaterThan(0));
            });

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false);

            Assert.That(reader.DocumentType, Is.EqualTo(SpreadsheetDocumentType.Workbook));

            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Workbook workbook = workbookPart!.Workbook;

            Sheets? sheets = workbook.Sheets;
            Assert.Multiple(() =>
            {
                Assert.That(sheets, Is.Not.Null);
                Assert.That(sheets!.Count(), Is.EqualTo(1));
            });
            Sheet sheet = (Sheet)sheets!.ChildElements[0];
            Assert.Multiple(() =>
            {
                Assert.That(sheet, Is.Not.Null);
                Assert.That(sheet.Name!.ToString(), Is.EqualTo("first"));
            });

            IEnumerable<Columns> columnsData = workbookPart.WorksheetParts.First().Worksheet.ChildElements.OfType<Columns>();
            Assert.Multiple(() =>
            {
                Assert.That(columnsData, Is.Not.Null);
                Assert.That(columnsData, Is.Empty);
            });

            IEnumerable<ConditionalFormatting> conditionalFormattings = GetConditionalFormatting(workbookPart.WorksheetParts.First().Worksheet);
            Assert.Multiple(() =>
            {
                Assert.That(conditionalFormattings, Has.Exactly(2).Matches<ConditionalFormatting>(c => string.Equals(c.SequenceOfReferences, "B1:B4", StringComparison.InvariantCultureIgnoreCase)));
                Assert.That(conditionalFormattings, Has.Exactly(1).Matches<ConditionalFormatting>
                    (c => c.ChildElements.First<ConditionalFormattingRule>()!.Operator! == ConditionalFormattingOperatorValues.LessThan
                        && c.ChildElements.First<ConditionalFormattingRule>()!.FirstChild!.InnerText == "5"
                        && c.ChildElements.First<ConditionalFormattingRule>()!.LastChild!.InnerText == "5"));
                Assert.That(conditionalFormattings, Has.Exactly(1).Matches<ConditionalFormatting>
                    (c => c.ChildElements.First<ConditionalFormattingRule>()!.Operator! == ConditionalFormattingOperatorValues.Between
                        && c.ChildElements.First<ConditionalFormattingRule>()!.FirstChild!.InnerText == "3"
                        && c.ChildElements.First<ConditionalFormattingRule>()!.LastChild!.InnerText == "7"));
            });

            IEnumerable<Row> rows = GetRows(workbookPart.WorksheetParts.First().Worksheet);
            Assert.Multiple(() =>
            {
                Assert.That(rows, Is.Not.Null);
                Assert.That(rows.Count(), Is.EqualTo(3));
            });

            int skipRows = 0;
            IEnumerable<Cell> cells = GetCells(rows.Skip(skipRows).First());
            Assert.Multiple(() =>
            {
                Assert.That(cells, Is.Not.Null);
                Assert.That(cells.Count(), Is.EqualTo(3));
                Assert.That(GetCellRealValue(cells.Skip(0).Take(1).First(), workbookPart), Is.EqualTo("a"));
                Assert.That(GetCellRealValue(cells.Skip(1).Take(1).First(), workbookPart), Is.EqualTo("b"));
                Assert.That(GetCellRealValue(cells.Skip(2).Take(1).First(), workbookPart), Is.EqualTo("c"));
            });
            skipRows++;

            cells = GetCells(rows.Skip(skipRows).First());
            Assert.Multiple(() =>
            {
                Assert.That(cells, Is.Not.Null);
                Assert.That(cells.Count(), Is.EqualTo(4));
                Assert.That(GetCellRealValue(cells.Skip(0).Take(1).First(), workbookPart), Is.EqualTo("1"));
                Assert.That(GetCellRealValue(cells.Skip(1).Take(1).First(), workbookPart), Is.EqualTo("2"));
                Assert.That(GetCellRealValue(cells.Skip(2).Take(1).First(), workbookPart), Is.EqualTo("30"));
                Assert.That(GetCellRealValue(cells.Skip(3).Take(1).First(), workbookPart), Is.EqualTo("40"));
            });
            skipRows++;

            cells = GetCells(rows.Skip(skipRows).First());
            Assert.Multiple(() =>
            {
                Assert.That(cells, Is.Not.Null);
                Assert.That(cells.Count(), Is.EqualTo(1));
                Assert.That(cells.Skip(0).Take(1).First().CellFormula!.Text, Is.EqualTo("SUM(A2:D2)"));
            });
        }

        [Test]
        public void InvalidStateRowOrSheet()
        {
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Multiple(() =>
                {
                    Assert.Throws<NoOpenSheetException>(() => writer.BeginRow());
                    Assert.Throws<NoOpenSheetException>(() => writer.BeginRow(1));
                    Assert.Throws<NoOpenRowException>(() => writer.EndRow());
                    Assert.Throws<NoOpenSheetException>(() => writer.CloseSheet());
                });
            }

            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                writer.CreateAndOpenSheet("abc");
                writer.BeginRow(2);
                writer.EndRow();
                Assert.Throws<OutOfOrderWritingException>(() => writer.BeginRow(1));
            }

            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                writer.CreateAndOpenSheet("abc");
                Assert.Throws<SheetAlreadyOpenException>(() => writer.CreateAndOpenSheet("opq"));
            }

            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                writer.CreateAndOpenSheet("abc");
                writer.BeginRow();
                Assert.Throws<RowAlreadyOpenException>(() => writer.BeginRow());
            }

            // -- //

            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Multiple(() =>
                {
                    Assert.Catch<InvalidOperationException>(() => writer.BeginRow());
                    Assert.Catch<InvalidOperationException>(() => writer.BeginRow(1));
                    Assert.Catch<InvalidOperationException>(() => writer.EndRow());
                    Assert.Catch<InvalidOperationException>(() => writer.CloseSheet());
                });
            }

            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                writer.CreateAndOpenSheet("abc");
                writer.BeginRow(2);
                writer.EndRow();
                Assert.Catch<InvalidOperationException>(() => writer.BeginRow(1));
            }

            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                writer.CreateAndOpenSheet("abc");
                Assert.Catch<InvalidOperationException>(() => writer.CreateAndOpenSheet("opq"));
            }

            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                writer.CreateAndOpenSheet("abc");
                writer.BeginRow();
                Assert.Catch<InvalidOperationException>(() => writer.BeginRow());
            }
        }

        [Test]
        public void InvalidStateCellThrowsNoOpenRowException()
        {
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Throws<NoOpenRowException>(() => writer.WriteTextCell("a"));
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                writer.CreateAndOpenSheet("name");
                Assert.Multiple(() =>
                {
                    Assert.Throws<NoOpenRowException>(() => writer.WriteTextCell("a"));
                    Assert.Throws<NoOpenRowException>(() => writer.WriteNumberCell(1f));
                    Assert.Throws<NoOpenRowException>(() => writer.WriteFormulaCell("SUM(A1:A2)"));
                });
            }
        }

        [Test]
        public void InvalidStateCellThrowsInvalidOperationException()
        {
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Catch<InvalidOperationException>(() => writer.WriteTextCell("a"));
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                writer.CreateAndOpenSheet("name");
                Assert.Multiple(() =>
                {
                    Assert.Catch<InvalidOperationException>(() => writer.WriteTextCell("a"));
                    Assert.Catch<InvalidOperationException>(() => writer.WriteNumberCell(1f));
                    Assert.Catch<InvalidOperationException>(() => writer.WriteFormulaCell("SUM(A1:A2)"));
                });
            }
        }

        [Test]
        public void SameResultsSharedStrings()
        {
            MemoryStream m1;
            MemoryStream m2;

            List<List<string>> strings =
            [
                ["Lorem ipsum", "dolor sit amet" ,"consectetur", "adipiscing elit", "Praesent at sapien", "id metus placerat" ,"ultricies", "a sed risus","Fusce finibus"],
                ["Lorem ipsum", "dolor sit amet", "Duis sodales finibus arcu", "porttitor", "accumsan", "finibus sapien", "ultricies", "a sed risus","Fusce finibus"],
                ["fermentum molestie", "parturient montes", "Lorem ipsum", "dolor sit amet" ,"eleifend", "urna", "laoreet libero", "id metus placerat" ,"justo convallis in"],
            ];

            using (BigExcelWriter writer1 = GetWriterStream(out m1))
            {
                writer1.CreateAndOpenSheet("s1");
                foreach (List<string> row in strings)
                {
                    writer1.WriteTextRow(row, useSharedStrings: true);
                }
                writer1.CloseSheet();
            }

            using (BigExcelWriter writer2 = GetWriterStream(out m2))
            {
                writer2.CreateAndOpenSheet("s1");
                foreach (List<string> row in strings)
                {
                    writer2.WriteTextRow(row, useSharedStrings: false);
                }
                writer2.CloseSheet();
            }

            using SpreadsheetDocument reader1 = SpreadsheetDocument.Open(m1, false);
            using SpreadsheetDocument reader2 = SpreadsheetDocument.Open(m2, false);

            WorkbookPart? workbookPart1 = reader1.WorkbookPart;
            Assert.That(workbookPart1, Is.Not.Null);
            IEnumerable<Row> rows1 = GetRows(workbookPart1!.WorksheetParts.First().Worksheet);
            WorkbookPart? workbookPart2 = reader2.WorkbookPart;
            Assert.That(workbookPart2, Is.Not.Null);
            IEnumerable<Row> rows2 = GetRows(workbookPart2!.WorksheetParts.First().Worksheet);

            Assert.Multiple(() =>
            {
                Assert.That(rows1, Is.Not.Null);
                Assert.That(rows1.Count(), Is.EqualTo(strings.Count));
                Assert.That(rows2, Is.Not.Null);
                Assert.That(rows2.Count(), Is.EqualTo(strings.Count));

                for (int i = 0; i < strings.Count; i++)
                {
                    IEnumerable<Cell> cells1 = GetCells(rows1.ElementAt(i));
                    IEnumerable<Cell> cells2 = GetCells(rows2.ElementAt(i));

                    Assert.That(cells1, Is.Not.Null);
                    Assert.That(cells2, Is.Not.Null);

                    Assert.That(cells1.Count(), Is.EqualTo(strings[i].Count));
                    Assert.That(cells2.Count(), Is.EqualTo(strings[i].Count));

                    for (int j = 0; j < strings[i].Count; j++)
                    {
                        Assert.That(GetCellRealValue(cells1.ElementAt(j), workbookPart1), Is.EqualTo(strings[i][j]));
                        Assert.That(GetCellRealValue(cells2.ElementAt(j), workbookPart2), Is.EqualTo(strings[i][j]));
                    }
                }
            });
        }

        [TestCase("text", "row")]
        [TestCase("number", "row")]
        [TestCase("formula", "row")]
        [TestCase("text", "cell")]
        [TestCase("number", "cell")]
        [TestCase("formula", "cell")]
        public void InvalidFormat(string @type, string rowOrCell)
        {
            using BigExcelWriter writer = GetWriterStream(out _);
            writer.CreateAndOpenSheet("a");
            switch (rowOrCell)
            {
                case "row":
                    switch (type)
                    {
                        case "text":
                            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteTextRow(["a"], -1));
                            break;
                        case "number":
                            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteNumberRow(new List<decimal> { 3m }, -1));
                            break;
                        case "formula":
                            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteFormulaRow(["a"], -1));
                            break;
                    }
                    break;
                case "cell":
                    writer.BeginRow();
                    switch (type)
                    {
                        case "text":
                            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteTextCell("a", -1));
                            break;
                        case "number":
                            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteNumberCell(3f, -1));
                            break;
                        case "formula":
                            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteFormulaCell("a", -1));
                            break;
                    }
                    writer.EndRow();
                    break;
            }
        }

        [Test]
        public void AutoFilter()
        {
            MemoryStream m1;

            List<List<string>> strings =
            [
                ["Lorem ipsum", "dolor sit amet" ,"consectetur", "adipiscing elit", "Praesent at sapien", "id metus placerat" ,"ultricies", "a sed risus","Fusce finibus"],
                ["Lorem ipsum", "dolor sit amet", "Duis sodales finibus arcu", "porttitor", "accumsan", "finibus sapien", "ultricies", "a sed risus","Fusce finibus"],
                ["fermentum molestie", "parturient montes", "Lorem ipsum", "dolor sit amet" ,"eleifend", "urna", "laoreet libero", "id metus placerat" ,"justo convallis in"],
            ];

            using (BigExcelWriter writer1 = GetWriterStream(out m1))
            {
                writer1.CreateAndOpenSheet("s1");
                foreach (List<string> row in strings)
                {
                    writer1.WriteTextRow(row, useSharedStrings: true);
                }
                writer1.AddAutofilter("A1:I1");
                writer1.CloseSheet();
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(m1, false);
            WorkbookPart workbookPart = reader.WorkbookPart!;
            Assert.That(workbookPart, Is.Not.Null);

            IEnumerable<AutoFilter> afs = workbookPart.WorksheetParts.First().Worksheet.ChildElements.OfType<AutoFilter>();
            Assert.That(afs, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(afs.Count(), Is.EqualTo(1));
                Assert.That(afs.First().Reference!.ToString(), Is.EqualTo("A1:I1"));
            });
        }

        [Test]
        public void AutoFilterErrorThrowsSheetAlreadyHasFilterException()
        {
            List<List<string>> strings =
            [
                ["Lorem ipsum", "dolor sit amet" ,"consectetur", "adipiscing elit", "Praesent at sapien", "id metus placerat" ,"ultricies", "a sed risus","Fusce finibus"],
                ["Lorem ipsum", "dolor sit amet", "Duis sodales finibus arcu", "porttitor", "accumsan", "finibus sapien", "ultricies", "a sed risus","Fusce finibus"],
                ["fermentum molestie", "parturient montes", "Lorem ipsum", "dolor sit amet" ,"eleifend", "urna", "laoreet libero", "id metus placerat" ,"justo convallis in"],
            ];

            using BigExcelWriter writer1 = GetWriterStream(out MemoryStream _);
            writer1.CreateAndOpenSheet("s1");
            foreach (List<string> row in strings)
            {
                writer1.WriteTextRow(row, useSharedStrings: true);
            }
            writer1.AddAutofilter("A1:I1");
            Assert.Throws<SheetAlreadyHasFilterException>(() => writer1.AddAutofilter("A1:J1"));
            writer1.CloseSheet();
        }

        [Test]
        public void AutoFilterErrorThrowsInvalidOperationException()
        {
            List<List<string>> strings =
            [
                ["Lorem ipsum", "dolor sit amet" ,"consectetur", "adipiscing elit", "Praesent at sapien", "id metus placerat" ,"ultricies", "a sed risus","Fusce finibus"],
                ["Lorem ipsum", "dolor sit amet", "Duis sodales finibus arcu", "porttitor", "accumsan", "finibus sapien", "ultricies", "a sed risus","Fusce finibus"],
                ["fermentum molestie", "parturient montes", "Lorem ipsum", "dolor sit amet" ,"eleifend", "urna", "laoreet libero", "id metus placerat" ,"justo convallis in"],
            ];

            using BigExcelWriter writer1 = GetWriterStream(out MemoryStream _);
            writer1.CreateAndOpenSheet("s1");
            foreach (List<string> row in strings)
            {
                writer1.WriteTextRow(row, useSharedStrings: true);
            }
            writer1.AddAutofilter("A1:I1");
            Assert.Catch<InvalidOperationException>(() => writer1.AddAutofilter("A1:J1"));
            writer1.CloseSheet();
        }

        [Test]
        public void ConditionalFormattingFormula()
        {
            MemoryStream memoryStream;
            using (BigExcelWriter writer = GetWriterStream(out memoryStream))
            {
                writer.CreateAndOpenSheet("a");
                for (int i = 0; i < 10; i++)
                {
                    writer.WriteNumberRow(new List<int> { i });
                }

                writer.AddConditionalFormattingFormula("A1:A20", "A1<5", 1);
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(memoryStream, false);
            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Assert.Multiple(() =>
            {
                Assert.That(workbookPart!.WorksheetParts.First().Worksheet.ChildElements.OfType<ConditionalFormatting>, Is.Not.Empty);
                IEnumerable<ConditionalFormatting> conditionalFormattings = workbookPart.WorksheetParts.First().Worksheet.ChildElements.OfType<ConditionalFormatting>();
                Assert.That(conditionalFormattings.Count(), Is.EqualTo(1));
                Assert.That(conditionalFormattings.First().SequenceOfReferences, Is.Not.Null);
                Assert.That(conditionalFormattings.First().SequenceOfReferences!.Items, Has.Count.EqualTo(1));
                Assert.That(conditionalFormattings.First().SequenceOfReferences!.Items.First().Value, Is.EqualTo("A1:A20"));
                Assert.That(conditionalFormattings.First().ChildElements.OfType<ConditionalFormattingRule>().Count(), Is.EqualTo(1));

                var rule = conditionalFormattings.First().ChildElements.OfType<ConditionalFormattingRule>().First();
                Assert.That(rule.Type, Is.Not.Null);
                Assert.That(rule.Type!.Value, Is.EqualTo(ConditionalFormatValues.Expression));
                Assert.That(rule.ChildElements, Has.Count.EqualTo(1));
                Assert.That(rule.ChildElements.OfType<Formula>().Count(), Is.EqualTo(1));
                Assert.That(rule.ChildElements.OfType<Formula>().First().Text, Is.EqualTo("A1<5"));
            });
        }

        [Test]
        public void ConditionalFormattingDuplicatedValues()
        {
            MemoryStream memoryStream;
            using (BigExcelWriter writer = GetWriterStream(out memoryStream))
            {
                writer.CreateAndOpenSheet("a");
                for (long i = 0; i < 10; i++)
                {
                    writer.WriteNumberRow(new List<long> { i });
                }

                writer.AddConditionalFormattingDuplicatedValues("A1:A20", 1);
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(memoryStream, false);
            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Assert.Multiple(() =>
            {
                Assert.That(workbookPart!.WorksheetParts.First().Worksheet.ChildElements.OfType<ConditionalFormatting>, Is.Not.Empty);
                IEnumerable<ConditionalFormatting> conditionalFormattings = workbookPart!.WorksheetParts.First().Worksheet.ChildElements.OfType<ConditionalFormatting>();
                Assert.That(conditionalFormattings.Count(), Is.EqualTo(1));
                Assert.That(conditionalFormattings.First().SequenceOfReferences, Is.Not.Null);
                Assert.That(conditionalFormattings.First().SequenceOfReferences!.Items, Has.Count.EqualTo(1));
                Assert.That(conditionalFormattings.First().SequenceOfReferences!.Items.First().Value, Is.EqualTo("A1:A20"));
                Assert.That(conditionalFormattings.First().ChildElements.OfType<ConditionalFormattingRule>().Count(), Is.EqualTo(1));

                var rule = conditionalFormattings.First().ChildElements.OfType<ConditionalFormattingRule>().First();
                Assert.That(rule.Type, Is.Not.Null);
                Assert.That(rule.Type!.Value, Is.EqualTo(ConditionalFormatValues.DuplicateValues));
                Assert.That(rule.ChildElements, Is.Empty);
            });
        }

        [TestCase(typeof(float), 1f, 10f)]
        [TestCase(typeof(float), 2f, 30f)]
        [TestCase(typeof(float), 3f, 30f)]
        [TestCase(typeof(double), 1d, 10d)]
        [TestCase(typeof(double), 2d, 30d)]
        [TestCase(typeof(double), 3d, 30d)]
        [TestCase(typeof(decimal), 1d, 10d)]
        [TestCase(typeof(decimal), 2d, 30d)]
        [TestCase(typeof(decimal), 3d, 30d)]
        public void DecimalValidator(Type type, object from, object to)
        {
            MemoryStream memoryStream;
            using (BigExcelWriter writer = GetWriterStream(out memoryStream))
            {
                writer.CreateAndOpenSheet("a");
                for (double i = 0; i < 10; i++)
                {
                    writer.WriteNumberRow(new List<double> { i });
                }

                switch (type)
                {
                    case Type t when t == typeof(float):
                        writer.AddDecimalValidator("A1:A20", (float)from, DataValidationOperatorValues.Between, secondOperand: (float)to);
                        break;
                    case Type t when t == typeof(double):
                        writer.AddDecimalValidator("A1:A20", (double)from, DataValidationOperatorValues.Between, secondOperand: (double)to);
                        break;
                    case Type t when t == typeof(decimal):
                        writer.AddDecimalValidator("A1:A20", Convert.ToDecimal(from), DataValidationOperatorValues.Between, secondOperand: Convert.ToDecimal(to));
                        break;
                    default:
                        Assert.Fail("Missing test case");
                        break;
                }
            }
            using SpreadsheetDocument reader = SpreadsheetDocument.Open(memoryStream, false);
            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Assert.Multiple(() =>
            {
                Assert.That(workbookPart!.WorksheetParts.First().Worksheet.ChildElements.OfType<DataValidations>, Is.Not.Empty);
                IEnumerable<DataValidations> dataValidations = workbookPart!.WorksheetParts.First().Worksheet.ChildElements.OfType<DataValidations>();
                Assert.That(dataValidations.Count(), Is.EqualTo(1));

                IEnumerable<DataValidation> dataValidationsE = dataValidations.First().ChildElements.OfType<DataValidation>();
                Assert.That(dataValidationsE.Count(), Is.EqualTo(1));

                DataValidation dataValidation = dataValidationsE.First();
                Assert.Multiple(() =>
                            {
                                Assert.That(dataValidation.Type!.Value, Is.EqualTo(DataValidationValues.Decimal));
                                Assert.That(dataValidation.Operator!.Value, Is.EqualTo(DataValidationOperatorValues.Between));
                                Assert.That(dataValidation.AllowBlank!.Value, Is.True);
                                Assert.That(dataValidation.ShowErrorMessage!.Value, Is.True);
                                Assert.That(dataValidation.ShowInputMessage!.Value, Is.True);
                                Assert.That(dataValidation.Formula1!.Text, Is.EqualTo(from.ToString()));
                                Assert.That(dataValidation.Formula2!.Text, Is.EqualTo(to.ToString()));
                            });
            });
        }

        [TestCase(typeof(float))]
        [TestCase(typeof(double))]
        [TestCase(typeof(decimal))]
        public void DecimalValidatorNoSecondOperand(Type type)
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            writer.CreateAndOpenSheet("a");
            for (uint i = 0; i < 10; i++)
            {
                writer.WriteNumberRow(new List<uint> { i });
            }

            switch (type)
            {
                case Type t when t == typeof(float):
                    Assert.Throws<ArgumentNullException>(() => writer.AddDecimalValidator("A1:A20", 1f, DataValidationOperatorValues.Between));
                    break;
                case Type t when t == typeof(double):
                    Assert.Throws<ArgumentNullException>(() => writer.AddDecimalValidator("A1:A20", 1d, DataValidationOperatorValues.Between));
                    break;
                case Type t when t == typeof(decimal):
                    Assert.Throws<ArgumentNullException>(() => writer.AddDecimalValidator("A1:A20", 1m, DataValidationOperatorValues.Between));
                    break;
                default:
                    Assert.Fail("Missing test case");
                    break;
            }
        }

        [TestCase(typeof(float))]
        [TestCase(typeof(double))]
        [TestCase(typeof(decimal))]
        public void DecimalValidationNoSheetThrowsNoOpenSheetException(Type type)
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            writer.CreateAndOpenSheet("a");
            for (decimal i = 0; i < 10; i++)
            {
                writer.WriteNumberRow(new List<decimal> { i });
            }
            writer.CloseSheet();

            switch (type)
            {
                case Type t when t == typeof(float):
                    Assert.Throws<NoOpenSheetException>(() => writer.AddDecimalValidator("A1:A20", 1f, DataValidationOperatorValues.Equal));
                    break;
                case Type t when t == typeof(double):
                    Assert.Throws<NoOpenSheetException>(() => writer.AddDecimalValidator("A1:A20", 1d, DataValidationOperatorValues.Equal));
                    break;
                case Type t when t == typeof(decimal):
                    Assert.Throws<NoOpenSheetException>(() => writer.AddDecimalValidator("A1:A20", 1m, DataValidationOperatorValues.Equal));
                    break;
                default:
                    Assert.Fail("Missing test case");
                    break;
            }
        }

        [TestCase(typeof(float))]
        [TestCase(typeof(double))]
        [TestCase(typeof(decimal))]
        public void DecimalValidationNoSheetThrowsInvalidOperationException(Type type)
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            writer.CreateAndOpenSheet("a");
            for (decimal i = 0; i < 10; i++)
            {
                writer.WriteNumberRow(new List<decimal> { i });
            }
            writer.CloseSheet();

            switch (type)
            {
                case Type t when t == typeof(float):
                    Assert.Catch<InvalidOperationException>(() => writer.AddDecimalValidator("A1:A20", 1f, DataValidationOperatorValues.Equal));
                    break;
                case Type t when t == typeof(double):
                    Assert.Catch<InvalidOperationException>(() => writer.AddDecimalValidator("A1:A20", 1d, DataValidationOperatorValues.Equal));
                    break;
                case Type t when t == typeof(decimal):
                    Assert.Catch<InvalidOperationException>(() => writer.AddDecimalValidator("A1:A20", 1m, DataValidationOperatorValues.Equal));
                    break;
                default:
                    Assert.Fail("Missing test case");
                    break;
            }
        }

        [TestCase(typeof(int))]
        [TestCase(typeof(uint))]
        [TestCase(typeof(long))]
        [TestCase(typeof(ulong))]
        public void IntegerValidator(Type type)
        {
            MemoryStream memoryStream;
            using (BigExcelWriter writer = GetWriterStream(out memoryStream))
            {
                writer.CreateAndOpenSheet("a");
                for (ulong i = 0; i < 10; i++)
                {
                    writer.WriteNumberRow(new List<ulong> { i });
                }

                switch (type)
                {
                    case Type t when t == typeof(int):
                        writer.AddIntegerValidator("A1:A20", 1, DataValidationOperatorValues.Equal);
                        break;
                    case Type t when t == typeof(uint):
                        writer.AddIntegerValidator("A1:A20", 1u, DataValidationOperatorValues.Equal);
                        break;
                    case Type t when t == typeof(long):
                        writer.AddIntegerValidator("A1:A20", 1L, DataValidationOperatorValues.Equal);
                        break;
                    case Type t when t == typeof(ulong):
                        writer.AddIntegerValidator("A1:A20", 1ul, DataValidationOperatorValues.Equal);
                        break;
                    default:
                        Assert.Fail("Missing test case");
                        break;
                }
            }
            using SpreadsheetDocument reader = SpreadsheetDocument.Open(memoryStream, false);
            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Assert.Multiple(() =>
            {
                Assert.That(workbookPart!.WorksheetParts.First().Worksheet.ChildElements.OfType<DataValidations>, Is.Not.Empty);
                IEnumerable<DataValidations> dataValidations = workbookPart.WorksheetParts.First().Worksheet.ChildElements.OfType<DataValidations>();
                Assert.That(dataValidations.Count(), Is.EqualTo(1));

                IEnumerable<DataValidation> dataValidationsE = dataValidations.First().ChildElements.OfType<DataValidation>();
                Assert.That(dataValidationsE.Count(), Is.EqualTo(1));

                DataValidation dataValidation = dataValidationsE.First();
                Assert.Multiple(() =>
                {
                    Assert.That(dataValidation.Type!.Value, Is.EqualTo(DataValidationValues.Whole));
                    Assert.That(dataValidation.Operator!.Value, Is.EqualTo(DataValidationOperatorValues.Equal));
                    Assert.That(dataValidation.AllowBlank!.Value, Is.True);
                    Assert.That(dataValidation.ShowErrorMessage!.Value, Is.True);
                    Assert.That(dataValidation.ShowInputMessage!.Value, Is.True);
                    Assert.That(dataValidation.Formula1!.Text, Is.EqualTo("1"));
                });
            });
        }

        [TestCase(typeof(int))]
        [TestCase(typeof(uint))]
        [TestCase(typeof(long))]
        [TestCase(typeof(ulong))]
        public void IntegerValidatorNoSecondOperand(Type type)
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            writer.CreateAndOpenSheet("a");
            for (byte i = 0; i < 10; i++)
            {
                writer.WriteNumberRow(new List<byte> { i });
            }

            switch (type)
            {
                case Type t when t == typeof(int):
                    Assert.Throws<ArgumentNullException>(() => writer.AddIntegerValidator("A1:A20", 1, DataValidationOperatorValues.Between));
                    break;
                case Type t when t == typeof(uint):
                    Assert.Throws<ArgumentNullException>(() => writer.AddIntegerValidator("A1:A20", 1u, DataValidationOperatorValues.Between));
                    break;
                case Type t when t == typeof(long):
                    Assert.Throws<ArgumentNullException>(() => writer.AddIntegerValidator("A1:A20", 1L, DataValidationOperatorValues.Between));
                    break;
                case Type t when t == typeof(ulong):
                    Assert.Throws<ArgumentNullException>(() => writer.AddIntegerValidator("A1:A20", 1ul, DataValidationOperatorValues.Between));
                    break;
                default:
                    Assert.Fail("Missing test case");
                    break;
            }
        }

        [TestCase(typeof(int))]
        [TestCase(typeof(uint))]
        [TestCase(typeof(long))]
        [TestCase(typeof(ulong))]
        public void IntegerValidationNoSheetThrowsNoOpenSheetException(Type type)
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            writer.CreateAndOpenSheet("a");
            for (sbyte i = 0; i < 10; i++)
            {
                writer.WriteNumberRow(new List<sbyte> { i });
            }
            writer.CloseSheet();

            switch (type)
            {
                case Type t when t == typeof(int):
                    Assert.Throws<NoOpenSheetException>(() => writer.AddIntegerValidator("A1:A20", 1, DataValidationOperatorValues.Equal));
                    break;
                case Type t when t == typeof(uint):
                    Assert.Throws<NoOpenSheetException>(() => writer.AddIntegerValidator("A1:A20", 1u, DataValidationOperatorValues.Equal));
                    break;
                case Type t when t == typeof(long):
                    Assert.Throws<NoOpenSheetException>(() => writer.AddIntegerValidator("A1:A20", 1L, DataValidationOperatorValues.Equal));
                    break;
                case Type t when t == typeof(ulong):
                    Assert.Throws<NoOpenSheetException>(() => writer.AddIntegerValidator("A1:A20", 1ul, DataValidationOperatorValues.Equal));
                    break;
                default:
                    Assert.Fail("Missing test case");
                    break;
            }
        }

        [TestCase(typeof(int))]
        [TestCase(typeof(uint))]
        [TestCase(typeof(long))]
        [TestCase(typeof(ulong))]
        public void IntegerValidationNoSheetThrowsInvalidOperationException(Type type)
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            writer.CreateAndOpenSheet("a");
            for (sbyte i = 0; i < 10; i++)
            {
                writer.WriteNumberRow(new List<sbyte> { i });
            }
            writer.CloseSheet();

            switch (type)
            {
                case Type t when t == typeof(int):
                    Assert.Catch<InvalidOperationException>(() => writer.AddIntegerValidator("A1:A20", 1, DataValidationOperatorValues.Equal));
                    break;
                case Type t when t == typeof(uint):
                    Assert.Catch<InvalidOperationException>(() => writer.AddIntegerValidator("A1:A20", 1u, DataValidationOperatorValues.Equal));
                    break;
                case Type t when t == typeof(long):
                    Assert.Catch<InvalidOperationException>(() => writer.AddIntegerValidator("A1:A20", 1L, DataValidationOperatorValues.Equal));
                    break;
                case Type t when t == typeof(ulong):
                    Assert.Catch<InvalidOperationException>(() => writer.AddIntegerValidator("A1:A20", 1ul, DataValidationOperatorValues.Equal));
                    break;
                default:
                    Assert.Fail("Missing test case");
                    break;
            }
        }

        [Test]
        public void ListValidator()
        {
            MemoryStream memoryStream;
            using (BigExcelWriter writer = GetWriterStream(out memoryStream))
            {
                writer.CreateAndOpenSheet("a");
                for (short i = 0; i < 10; i++)
                {
                    writer.WriteNumberRow(new List<short> { i });
                }

                writer.AddListValidator("A1:A20", "B1:B4");
            }
            using SpreadsheetDocument reader = SpreadsheetDocument.Open(memoryStream, false);
            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Assert.Multiple(() =>
            {
                Assert.That(workbookPart!.WorksheetParts.First().Worksheet.ChildElements.OfType<DataValidations>, Is.Not.Empty);
                IEnumerable<DataValidations> dataValidations = workbookPart!.WorksheetParts.First().Worksheet.ChildElements.OfType<DataValidations>();
                Assert.That(dataValidations.Count(), Is.EqualTo(1));

                IEnumerable<DataValidation> dataValidationsE = dataValidations.First().ChildElements.OfType<DataValidation>();
                Assert.That(dataValidationsE.Count(), Is.EqualTo(1));

                DataValidation dataValidation = dataValidationsE.First();
                Assert.Multiple(() =>
                {
                    Assert.That(dataValidation.Type!.Value, Is.EqualTo(DataValidationValues.List));
                    Assert.That(dataValidation.Operator!.Value, Is.EqualTo(DataValidationOperatorValues.Equal));
                    Assert.That(dataValidation.AllowBlank!.Value, Is.True);
                    Assert.That(dataValidation.ShowErrorMessage!.Value, Is.True);
                    Assert.That(dataValidation.ShowInputMessage!.Value, Is.True);
                    Assert.That(dataValidation.Formula1!.Text, Is.EqualTo("B1:B4"));
                });
            });
        }

        [Test]
        public void MergedCells()
        {
            MemoryStream memoryStream;
            using (BigExcelWriter writer = GetWriterStream(out memoryStream))
            {
                writer.CreateAndOpenSheet("a");
                writer.MergeCells("a");
                writer.MergeCells("b2:d7");
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(memoryStream, false);
            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Assert.Multiple(() =>
            {
                Assert.That(workbookPart!.WorksheetParts.First().Worksheet.ChildElements.OfType<MergeCells>, Is.Not.Empty);
                IEnumerable<MergeCells> mergedCellsElement = workbookPart!.WorksheetParts.First().Worksheet.ChildElements.OfType<MergeCells>();
                Assert.Multiple(() =>
                {
                    Assert.That(mergedCellsElement, Is.Not.Null);
                    Assert.That(mergedCellsElement.Count, Is.EqualTo(1));
                });

                IEnumerable<MergeCell> mergedCellElements = mergedCellsElement.First().ChildElements.OfType<MergeCell>();
                Assert.Multiple(() =>
                {
                    Assert.That(mergedCellElements, Is.Not.Null);
                    Assert.That(mergedCellElements.Count, Is.EqualTo(2));
                    Assert.That(mergedCellElements, Has.Exactly(1).Matches<MergeCell>(mce => mce.Reference!.Value!.Equals("A:A", StringComparison.InvariantCultureIgnoreCase)));
                    Assert.That(mergedCellElements, Has.Exactly(1).Matches<MergeCell>(mce => mce.Reference!.Value!.Equals("B2:D7", StringComparison.InvariantCultureIgnoreCase)));
                });
            });
        }

        [Test]
        public void MergedCellsOverlappingRangesThrowsOverlappingRangesException()
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            writer.CreateAndOpenSheet("a");
            writer.MergeCells("a1:c7");
            Assert.Throws<OverlappingRangesException>(() => writer.MergeCells("b2:b3"));
        }

        [Test]
        public void MergedCellsOverlappingRangesThrowsInvalidOperationException()
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            writer.CreateAndOpenSheet("a");
            writer.MergeCells("a1:c7");
            Assert.Catch<InvalidOperationException>(() => writer.MergeCells("b2:b3"));
        }

        [Test]
        public void MergedCellsNoSheetThrowsNoOpenSheetException()
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            Assert.Throws<NoOpenSheetException>(() => writer.MergeCells("b2:b3"));
        }

        [Test]
        public void MergedCellsNoSheetThrowsInvalidOperationException()
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            Assert.Catch<InvalidOperationException>(() => writer.MergeCells("b2:b3"));
        }

        [Test]
        public void PageLayout()
        {
            MemoryStream memoryStream;
            using (BigExcelWriter writer = GetWriterStream(out memoryStream))
            {
                writer.CreateAndOpenSheet("a");
                writer.CloseSheet();

                writer.CreateAndOpenSheet("hideGrid");
                writer.ShowGridLinesInCurrentSheet = false;
                writer.CloseSheet();

                writer.CreateAndOpenSheet("hideAndPrintGrid");
                writer.ShowGridLinesInCurrentSheet = false;
                writer.PrintGridLinesInCurrentSheet = true;
                writer.CloseSheet();

                writer.CreateAndOpenSheet("hideAndPrintHead");
                writer.ShowRowAndColumnHeadingsInCurrentSheet = false;
                writer.PrintRowAndColumnHeadingsInCurrentSheet = true;
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(memoryStream, false);
            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.Multiple(() =>
            {
                Assert.That(workbookPart, Is.Not.Null);
                Assert.That(workbookPart!.WorksheetParts.Count, Is.EqualTo(4));
            });

            Assert.Multiple(() =>
            {
                Worksheet worksheet = workbookPart!.WorksheetParts.ElementAt(0).Worksheet;
                IEnumerable<SheetViews> sheetViewsParts = worksheet.ChildElements.OfType<SheetViews>();
                Assert.That(sheetViewsParts, Is.Empty);

                var printOptionsPart = worksheet.ChildElements.OfType<PrintOptions>();
                Assert.That(printOptionsPart, Is.Empty);
            });

            Assert.Multiple(() =>
            {
                Worksheet worksheet = workbookPart!.WorksheetParts.ElementAt(1).Worksheet;
                IEnumerable<SheetViews> sheetViewsParts = worksheet.ChildElements.OfType<SheetViews>();
                Assert.That(sheetViewsParts.Count, Is.EqualTo(1));
                SheetViews sheetViews = sheetViewsParts.ElementAt(0);
                Assert.That(sheetViews.ChildElements.OfType<SheetView>().Count, Is.EqualTo(1));
                SheetView sheetView = (SheetView)sheetViews.First();
                Assert.That(sheetView, Is.Not.Null);
                Assert.That(sheetView.ShowGridLines, Is.Not.Null);
                Assert.That(sheetView.ShowGridLines!.Value, Is.False);
                Assert.That(sheetView.ShowRowColHeaders, Is.Null);
            });

            Assert.Multiple(() =>
            {
                Worksheet worksheet = workbookPart!.WorksheetParts.ElementAt(2).Worksheet;
                IEnumerable<SheetViews> sheetViewsParts = worksheet.ChildElements.OfType<SheetViews>();
                Assert.That(sheetViewsParts.Count, Is.EqualTo(1));
                SheetViews sheetViews = sheetViewsParts.ElementAt(0);
                Assert.That(sheetViews.ChildElements.OfType<SheetView>().Count, Is.EqualTo(1));
                SheetView sheetView = (SheetView)sheetViews.First();
                Assert.That(sheetView, Is.Not.Null);
                Assert.That(sheetView.ShowGridLines, Is.Not.Null);
                Assert.That(sheetView.ShowGridLines!.Value, Is.False);
                Assert.That(sheetView.ShowRowColHeaders, Is.Null);

                var printOptionsPart = worksheet.ChildElements.OfType<PrintOptions>();
                Assert.That(printOptionsPart.Count, Is.EqualTo(1));
                var printOptions = printOptionsPart.First();
                Assert.That(printOptions, Is.Not.Null);
                Assert.That(printOptions.GridLines, Is.Not.Null);
                Assert.That(printOptions.GridLines!.Value, Is.True);
                Assert.That(printOptions.Headings, Is.Null);
            });

            Assert.Multiple(() =>
            {
                Worksheet worksheet = workbookPart!.WorksheetParts.ElementAt(3).Worksheet;
                IEnumerable<SheetViews> sheetViewsParts = worksheet.ChildElements.OfType<SheetViews>();
                Assert.That(sheetViewsParts.Count, Is.EqualTo(1));
                SheetViews sheetViews = sheetViewsParts.ElementAt(0);
                Assert.That(sheetViews.ChildElements.OfType<SheetView>().Count, Is.EqualTo(1));
                SheetView sheetView = (SheetView)sheetViews.First();
                Assert.That(sheetView.ShowGridLines, Is.Null);
                Assert.That(sheetView.ShowRowColHeaders, Is.Not.Null);
                Assert.That(sheetView.ShowRowColHeaders!.Value, Is.False);

                var printOptionsPart = worksheet.ChildElements.OfType<PrintOptions>();
                Assert.That(printOptionsPart.Count, Is.EqualTo(1));
                var printOptions = printOptionsPart.First();
                Assert.That(printOptions.GridLines, Is.Null);
                Assert.That(printOptions.Headings, Is.Not.Null);
                Assert.That(printOptions.Headings!.Value, Is.True);
            });
        }

        [Test]
        public void PageLayoutReturnsToDefault()
        {
            using BigExcelWriter writer = GetWriterStream(out _);

            writer.CreateAndOpenSheet("a");

            Assert.Multiple(() =>
            {
                Assert.That(writer.ShowGridLinesInCurrentSheet, Is.True);
                Assert.That(writer.ShowRowAndColumnHeadingsInCurrentSheet, Is.True);
                Assert.That(writer.PrintGridLinesInCurrentSheet, Is.False);
                Assert.That(writer.PrintRowAndColumnHeadingsInCurrentSheet, Is.False);
            });

            Assert.That(() =>
            {
                writer.ShowGridLinesInCurrentSheet = false;
                writer.ShowRowAndColumnHeadingsInCurrentSheet = false;
                writer.PrintGridLinesInCurrentSheet = true;
                writer.PrintRowAndColumnHeadingsInCurrentSheet = true;
            }
            , Throws.Nothing);

            Assert.Multiple(() =>
            {
                Assert.That(writer.ShowGridLinesInCurrentSheet, Is.False);
                Assert.That(writer.ShowRowAndColumnHeadingsInCurrentSheet, Is.False);
                Assert.That(writer.PrintGridLinesInCurrentSheet, Is.True);
                Assert.That(writer.PrintRowAndColumnHeadingsInCurrentSheet, Is.True);
            });
        }

        [Test]
        public void PageLayoutInvalidContextThrowsNoOpenSheetException()
        {
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Throws<NoOpenSheetException>(() => writer.ShowGridLinesInCurrentSheet = false);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Throws<NoOpenSheetException>(() => writer.ShowRowAndColumnHeadingsInCurrentSheet = false);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Throws<NoOpenSheetException>(() => writer.PrintGridLinesInCurrentSheet = true);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Throws<NoOpenSheetException>(() => writer.PrintRowAndColumnHeadingsInCurrentSheet = true);
            }

            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Throws<NoOpenSheetException>(() => _ = writer.ShowGridLinesInCurrentSheet);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Throws<NoOpenSheetException>(() => _ = writer.ShowRowAndColumnHeadingsInCurrentSheet);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Throws<NoOpenSheetException>(() => _ = writer.PrintGridLinesInCurrentSheet);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Throws<NoOpenSheetException>(() => _ = writer.PrintRowAndColumnHeadingsInCurrentSheet);
            }
        }

        [Test]
        public void PageLayoutInvalidContextThrowsInvalidOperationException()
        {
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Catch<InvalidOperationException>(() => writer.ShowGridLinesInCurrentSheet = false);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Catch<InvalidOperationException>(() => writer.ShowRowAndColumnHeadingsInCurrentSheet = false);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Catch<InvalidOperationException>(() => writer.PrintGridLinesInCurrentSheet = true);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Catch<InvalidOperationException>(() => writer.PrintRowAndColumnHeadingsInCurrentSheet = true);
            }

            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Catch<InvalidOperationException>(() => _ = writer.ShowGridLinesInCurrentSheet);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Catch<InvalidOperationException>(() => _ = writer.ShowRowAndColumnHeadingsInCurrentSheet);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Catch<InvalidOperationException>(() => _ = writer.PrintGridLinesInCurrentSheet);
            }
            using (BigExcelWriter writer = GetWriterStream(out _))
            {
                Assert.Catch<InvalidOperationException>(() => _ = writer.PrintRowAndColumnHeadingsInCurrentSheet);
            }
        }

        [TestCase("")]
        [TestCase(null)]
        public void SheetNameEmptyThrowsSheetNameCannotBeEmptyException(string? sheetName)
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            Assert.Throws<SheetNameCannotBeEmptyException>(() => writer.CreateAndOpenSheet(sheetName));
        }

        [TestCase("")]
        [TestCase(null)]
        public void SheetNameEmptyThrowsInvalidOperationException(string? sheetName)
        {
            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            Assert.Catch<InvalidOperationException>(() => writer.CreateAndOpenSheet(sheetName));
        }

        [TestCase("a", "a")]
        [TestCase("a", "A")]
        [TestCase("A", "a")]
        [TestCase("A", "A")]
        [TestCase("b", "B")]
        [TestCase("AB", "ab")]
        [TestCase("aB", "Ab")]
        [TestCase("Ab", "aB")]
        public void SheetNameRepeatedThrowsSheetWithSameNameAlreadyExistsException(string a, string b)
        {
            if (string.Compare(a, b, true) != 0)
            {
                Assert.Fail("Precondition failed. a and b must be equal except for case.");
            }

            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            writer.CreateAndOpenSheet("a");
            writer.CloseSheet();
            Assert.Throws<SheetWithSameNameAlreadyExistsException>(() => writer.CreateAndOpenSheet("a"));
        }

        [TestCase("a", "a")]
        [TestCase("a", "A")]
        [TestCase("A", "a")]
        [TestCase("A", "A")]
        [TestCase("b", "B")]
        [TestCase("AB", "ab")]
        [TestCase("aB", "Ab")]
        [TestCase("Ab", "aB")]
        public void SheetNameRepeatedThrowsInvalidOperationException(string a, string b)
        {
            if (string.Compare(a, b, true) != 0)
            {
                Assert.Fail("Precondition failed. a and b must be equal except for case.");
            }

            using BigExcelWriter writer = GetWriterStream(out MemoryStream _);
            writer.CreateAndOpenSheet("a");
            writer.CloseSheet();
            Assert.Catch<InvalidOperationException>(() => writer.CreateAndOpenSheet("a"));
        }

        [Test]
        public void FromObjectTest()
        {
            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream))
            {

                List<ExampleModel> data =
                [
                    new ExampleModel
                    {
                        Id = 1,
                        Name = "John Doe",
                        Position = "Software Engineer",
                        Description = "Develops software solutions.",
                        CreatedAt = new DateTime(2022, 1, 15),
                        Salary = 6000.50m,
                        Sale = 15000.75m,
                        Secret = "TopSecret",
                    },
                    new ExampleModel
                    {
                        Id = 2,
                        Name = "Jane Smith",
                        Position = "Project Manager",
                        Description = "Manages projects and teams.",
                        CreatedAt = new DateTime(2021, 11, 30),
                        Salary = 7500.00m,
                        Sale = 20000.00m,
                        Secret = "Classified",
                    }
                ];

                writer.CreateSheetFromObject(data, "example");
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false);


            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Workbook workbook = workbookPart!.Workbook;

            Sheets? sheets = workbook.Sheets;
            Assert.Multiple(() =>
            {
                Assert.That(sheets, Is.Not.Null);
                Assert.That(sheets!.Count(), Is.EqualTo(1));
            });
            Sheet sheet = (Sheet)sheets!.ChildElements[0];
            Assert.Multiple(() =>
            {
                Assert.That(sheet, Is.Not.Null);
                Assert.That(sheet.Name!.ToString(), Is.EqualTo("example"));
            });
        }
    }
}
