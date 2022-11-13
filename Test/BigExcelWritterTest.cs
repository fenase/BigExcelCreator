using BigExcelCreator;
using BigExcelCreator.Ranges;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static Test.TestHelperMethods;

namespace Test
{
    internal class BigExcelWriterTest
    {
        string DirectoryPath { get; set; }


        [SetUp]
        public void Setup()
        {
            DirectoryPath = TestContext.CurrentContext.WorkDirectory + @"\excelOut";
            Directory.CreateDirectory(DirectoryPath);
            DirectoryAssert.Exists(DirectoryPath);
        }



        [TearDown]
        public void Teardown()
        {
            new DirectoryInfo(DirectoryPath).Delete(true);
        }


        [Test]
        public void FileExistsAfterCreation()
        {
            string path = Path.Combine(DirectoryPath, $"{Guid.NewGuid()}.xlsx");
            using (BigExcelWriter writer = new BigExcelWriter(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                // do nothing
            }
            FileAssert.Exists(path);
        }


        [Test]
        public void ValidFile()
        {
            string path = Path.Combine(DirectoryPath, $"{Guid.NewGuid()}.xlsx");
            using (BigExcelWriter writer = new BigExcelWriter(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                writer.CreateAndOpenSheet("first");
                writer.CloseSheet();
            }
            FileAssert.Exists(path);

            using (SpreadsheetDocument reader = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart? workbookPart = reader.WorkbookPart;
                Assert.That(workbookPart, Is.Not.Null);

                Workbook workbook = workbookPart.Workbook;

                Sheets? sheets = workbook.Sheets;
                Assert.Multiple(() =>
                {
                    Assert.That(sheets, Is.Not.Null);
                    Assert.That(sheets!.Count(), Is.EqualTo(1));
                });
                Sheet sheet = (Sheet)sheets!.ChildElements.First();
                Assert.Multiple(() =>
                {
                    Assert.That(sheet, Is.Not.Null);
                    Assert.That(sheet.Name!.ToString(), Is.EqualTo("first"));
                });
            }
        }


        [Test]
        public void ValidStream()
        {
            MemoryStream stream = new MemoryStream();
            using (BigExcelWriter writer = new BigExcelWriter(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                writer.CreateAndOpenSheet("first");
                writer.CloseSheet();
            }
            Assert.Multiple(() =>
            {

                Assert.That(stream.Position, Is.EqualTo(0));
                Assert.That(stream, Has.Length.GreaterThan(0));
            });

            using (SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart? workbookPart = reader.WorkbookPart;
                Assert.That(workbookPart, Is.Not.Null);

                Workbook workbook = workbookPart.Workbook;

                Sheets? sheets = workbook.Sheets;
                Assert.Multiple(() =>
                {
                    Assert.That(sheets, Is.Not.Null);
                    Assert.That(sheets!.Count(), Is.EqualTo(1));
                });
                Sheet sheet = (Sheet)sheets!.ChildElements.First();
                Assert.Multiple(() =>
                {
                    Assert.That(sheet, Is.Not.Null);
                    Assert.That(sheet.Name!.ToString(), Is.EqualTo("first"));
                });
            }
        }


        [Test]
        public void ValidContent()
        {
            MemoryStream stream = new MemoryStream();
            using (BigExcelWriter writer = new BigExcelWriter(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                writer.CreateAndOpenSheet("first");
                writer.WriteTextRow(new[] { "a", "b", "c" });
                writer.WriteNumberRow(new[] { 1f, 2f, 30f, 40f });
                writer.WriteFormulaRow(new[] { "SUM(A2:D2)" });
                writer.CloseSheet();
            }

            Assert.Multiple(() =>
            {
                Assert.That(stream.Position, Is.EqualTo(0));
                Assert.That(stream, Has.Length.GreaterThan(0));
            });

            using (SpreadsheetDocument reader = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart? workbookPart = reader.WorkbookPart;
                Assert.That(workbookPart, Is.Not.Null);

                Workbook workbook = workbookPart.Workbook;

                Sheets? sheets = workbook.Sheets;
                Assert.Multiple(() =>
                {
                    Assert.That(sheets, Is.Not.Null);
                    Assert.That(sheets!.Count(), Is.EqualTo(1));
                });
                Sheet sheet = (Sheet)sheets!.ChildElements.First();
                Assert.Multiple(() =>
                {
                    Assert.That(sheet, Is.Not.Null);
                    Assert.That(sheet.Name!.ToString(), Is.EqualTo("first"));
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
        }


        [Test]
        public void InvalidStateRowOrSheet()
        {
            using (BigExcelWriter writer = GetwriterStream(out _))
            {
                Assert.Multiple(() =>
                {
                    Assert.Throws<InvalidOperationException>(() => writer.BeginRow());
                    Assert.Throws<InvalidOperationException>(() => writer.BeginRow(1));
                    Assert.Throws<InvalidOperationException>(() => writer.EndRow());
                    Assert.Throws<InvalidOperationException>(() => writer.CloseSheet());
                });
            }

            using (BigExcelWriter writer = GetwriterStream(out _))
            {
                writer.CreateAndOpenSheet("abc");
                writer.BeginRow(2);
                writer.EndRow();
                Assert.Throws<InvalidOperationException>(() => writer.BeginRow(1));
            }

            using (BigExcelWriter writer = GetwriterStream(out _))
            {
                writer.CreateAndOpenSheet("abc");
                Assert.Throws<InvalidOperationException>(() => writer.CreateAndOpenSheet("opq"));
            }
        }

        [Test]
        public void InvalidStateCell()
        {
            using (BigExcelWriter writer = GetwriterStream(out _))
            {
                Assert.Throws<InvalidOperationException>(() => writer.WriteTextCell("a"));
            }
            using (BigExcelWriter writer = GetwriterStream(out _))
            {
                writer.CreateAndOpenSheet("name");
                Assert.Multiple(() =>
                {
                    Assert.Throws<InvalidOperationException>(() => writer.WriteTextCell("a"));
                    Assert.Throws<InvalidOperationException>(() => writer.WriteNumberCell(1f));
                    Assert.Throws<InvalidOperationException>(() => writer.WriteFormulaCell("SUM(A1:A2)"));
                });
            }
        }

        [Test]
        public void SameResultsSharedStrings()
        {
            MemoryStream m1;
            MemoryStream m2;

            List<List<string>> strings = new List<List<string>>
            {
                new List<string>{ "Lorem ipsum", "dolor sit amet" ,"consectetur", "adipiscing elit", "Praesent at sapien", "id metus placerat" ,"ultricies", "a sed risus","Fusce finibus"},
                new List<string>{ "Lorem ipsum", "dolor sit amet", "Duis sodales finibus arcu", "porttitor", "accumsan", "finibus sapien", "ultricies", "a sed risus","Fusce finibus"},
                new List<string>{ "fermentum molestie", "parturient montes", "Lorem ipsum", "dolor sit amet" ,"eleifend", "urna", "laoreet libero", "id metus placerat" ,"justo convallis in"},
            };

            using (BigExcelWriter writer1 = GetwriterStream(out m1))
            {
                writer1.CreateAndOpenSheet("s1");
                foreach (List<string> row in strings)
                {
                    writer1.WriteTextRow(row, useSharedStrings: true);
                }
                writer1.CloseSheet();
            }

            using (BigExcelWriter writer2 = GetwriterStream(out m2))
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
            IEnumerable<Row> rows1 = GetRows(workbookPart1.WorksheetParts.First().Worksheet);
            WorkbookPart? workbookPart2 = reader2.WorkbookPart;
            Assert.That(workbookPart2, Is.Not.Null);
            IEnumerable<Row> rows2 = GetRows(workbookPart2.WorksheetParts.First().Worksheet);

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
            using BigExcelWriter writer = GetwriterStream(out _);
            writer.CreateAndOpenSheet("a");
            switch (rowOrCell)
            {
                case "row":
                    switch (type)
                    {
                        case "text":
                            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteTextRow(new[] { "a" }, -1));
                            break;
                        case "number":
                            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteNumberRow(new[] { 3f }, -1));
                            break;
                        case "formula":
                            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteFormulaRow(new[] { "a" }, -1));
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
        public void ConditionalFormattingFormula()
        {
            MemoryStream memoryStream;
            using (BigExcelWriter writer = GetwriterStream(out memoryStream))
            {
                writer.CreateAndOpenSheet("a");
                for (int i = 0; i < 10; i++)
                {
                    writer.WriteNumberRow(new List<float> { i });
                }

                writer.AddConditionalFormattingFormula("A1:A20", "A1<5", 1);
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(memoryStream, false);
            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Assert.Multiple(() =>
            {
                Assert.That(workbookPart.WorksheetParts.First().Worksheet.ChildElements.OfType<ConditionalFormatting>, Is.Not.Empty);
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
            using (BigExcelWriter writer = GetwriterStream(out memoryStream))
            {
                writer.CreateAndOpenSheet("a");
                for (int i = 0; i < 10; i++)
                {
                    writer.WriteNumberRow(new List<float> { i });
                }

                writer.AddConditionalFormattingDuplicatedValues("A1:A20", 1);
            }

            using SpreadsheetDocument reader = SpreadsheetDocument.Open(memoryStream, false);
            WorkbookPart? workbookPart = reader.WorkbookPart;
            Assert.That(workbookPart, Is.Not.Null);

            Assert.Multiple(() =>
            {
                Assert.That(workbookPart.WorksheetParts.First().Worksheet.ChildElements.OfType<ConditionalFormatting>, Is.Not.Empty);
                IEnumerable<ConditionalFormatting> conditionalFormattings = workbookPart.WorksheetParts.First().Worksheet.ChildElements.OfType<ConditionalFormatting>();
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

        [Test]
        public void MergedCells()
        {
            MemoryStream memoryStream;
            using (BigExcelWriter writer = GetwriterStream(out memoryStream))
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
                Assert.That(workbookPart.WorksheetParts.First().Worksheet.ChildElements.OfType<MergeCells>, Is.Not.Empty);
                IEnumerable<MergeCells> mergedCellsElement = workbookPart.WorksheetParts.First().Worksheet.ChildElements.OfType<MergeCells>();
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
                    Assert.That(mergedCellElements, Has.Exactly(1).Matches<MergeCell>(mce => mce.Reference.Value.Equals("A:A", StringComparison.InvariantCultureIgnoreCase)));
                    Assert.That(mergedCellElements, Has.Exactly(1).Matches<MergeCell>(mce => mce.Reference.Value.Equals("B2:D7", StringComparison.InvariantCultureIgnoreCase)));
                });
            });
        }

        [Test]
        public void MergedCellsOverlappingRanges()
        {
            using BigExcelWriter writer = GetwriterStream(out MemoryStream memoryStream);
            writer.CreateAndOpenSheet("a");
            writer.MergeCells("a1:c7");
            Assert.Throws<OverlappingRangesException>(() => writer.MergeCells("b2:b3"));
        }

        [Test]
        public void MergedCellsNoSheet()
        {
            using BigExcelWriter writer = GetwriterStream(out MemoryStream memoryStream);
            Assert.Throws<InvalidOperationException>(() => writer.MergeCells("b2:b3"));
        }
    }
}
