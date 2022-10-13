using BigExcelCreator;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Test
{
    internal class BigExcelWritterTest
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
            string path = Path.Combine(DirectoryPath, "creationTest.xlsx");
            using (BigExcelWritter writter = new BigExcelWritter(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                // do nothing
            }
            FileAssert.Exists(path);
        }


        [Test]
        public void ValidFile()
        {
            string path = Path.Combine(DirectoryPath, "ValidFile.xlsx");
            using (BigExcelWritter writter = new BigExcelWritter(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                writter.CreateAndOpenSheet("first");
                writter.CloseSheet();
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
            using (BigExcelWritter writter = new BigExcelWritter(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                writter.CreateAndOpenSheet("first");
                writter.CloseSheet();
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
            using (BigExcelWritter writter = new BigExcelWritter(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                writter.CreateAndOpenSheet("first");
                writter.WriteTextRow(new[] { "a", "b", "c" });
                writter.WriteNumberRow(new[] { 1f, 2f, 30f, 40f });
                writter.WriteFormulaRow(new[] { "SUM(A2:D2)" });
                writter.CloseSheet();
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
                IEnumerable<Cell> cells = GetCells(rows.Skip(skipRows++).First());
                Assert.Multiple(() =>
                {
                    Assert.That(cells, Is.Not.Null);
                    Assert.That(cells.Count(), Is.EqualTo(3));
                    Assert.That(GetCellRealValue(cells.Skip(0).Take(1).First(), workbookPart), Is.EqualTo("a"));
                    Assert.That(GetCellRealValue(cells.Skip(1).Take(1).First(), workbookPart), Is.EqualTo("b"));
                    Assert.That(GetCellRealValue(cells.Skip(2).Take(1).First(), workbookPart), Is.EqualTo("c"));
                });

                cells = GetCells(rows.Skip(skipRows++).First());
                Assert.Multiple(() =>
                {
                    Assert.That(cells, Is.Not.Null);
                    Assert.That(cells.Count(), Is.EqualTo(4));
                    Assert.That(GetCellRealValue(cells.Skip(0).Take(1).First(), workbookPart), Is.EqualTo("1"));
                    Assert.That(GetCellRealValue(cells.Skip(1).Take(1).First(), workbookPart), Is.EqualTo("2"));
                    Assert.That(GetCellRealValue(cells.Skip(2).Take(1).First(), workbookPart), Is.EqualTo("30"));
                    Assert.That(GetCellRealValue(cells.Skip(3).Take(1).First(), workbookPart), Is.EqualTo("40"));
                });

                cells = GetCells(rows.Skip(skipRows++).First());
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
            using (BigExcelWritter writter = GetWritterStream(out _))
            {
                Assert.Multiple(() =>
                {
                    Assert.Throws<InvalidOperationException>(() => writter.BeginRow());
                    Assert.Throws<InvalidOperationException>(() => writter.BeginRow(1));
                    Assert.Throws<InvalidOperationException>(() => writter.EndRow());
                    Assert.Throws<InvalidOperationException>(() => writter.CloseSheet());
                });
            }

            using (BigExcelWritter writter = GetWritterStream(out _))
            {
                writter.CreateAndOpenSheet("abc");
                writter.BeginRow(2);
                writter.EndRow();
                Assert.Throws<InvalidOperationException>(() => writter.BeginRow(1));
            }

            using (BigExcelWritter writter = GetWritterStream(out _))
            {
                writter.CreateAndOpenSheet("abc");
                Assert.Throws<InvalidOperationException>(() => writter.CreateAndOpenSheet("opq"));
            }
        }

        [Test]
        public void InvalidStateCell()
        {
            using (BigExcelWritter writter = GetWritterStream(out _))
            {
                Assert.Throws<InvalidOperationException>(() => writter.WriteTextCell("a"));
            }
            using (BigExcelWritter writter = GetWritterStream(out _))
            {
                writter.CreateAndOpenSheet("name");
                Assert.Multiple(() =>
                {
                    Assert.Throws<InvalidOperationException>(() => writter.WriteTextCell("a"));
                    Assert.Throws<InvalidOperationException>(() => writter.WriteNumberCell(1f));
                    Assert.Throws<InvalidOperationException>(() => writter.WriteFormulaCell("SUM(A1:A2)"));
                });
            }
        }

        [Test]
        public void SameResultsSharedStrings()
        {
            MemoryStream m1 = new MemoryStream();
            MemoryStream m2 = new MemoryStream();

            List<List<string>> strings = new List<List<string>>
            {
                new List<string>{ "Lorem ipsum", "dolor sit amet" ,"consectetur", "adipiscing elit", "Praesent at sapien", "id metus placerat" ,"ultricies", "a sed risus","Fusce finibus"},
                new List<string>{ "Lorem ipsum", "dolor sit amet", "Duis sodales finibus arcu", "porttitor", "accumsan", "finibus sapien", "ultricies", "a sed risus","Fusce finibus"},
                new List<string>{ "fermentum molestie", "parturient montes", "Lorem ipsum", "dolor sit amet" ,"eleifend", "urna", "laoreet libero", "id metus placerat" ,"justo convallis in"},
            };

            using (BigExcelWritter writter1 = new BigExcelWritter(m1, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                writter1.CreateAndOpenSheet("s1");
                foreach (List<string> row in strings)
                {
                    writter1.WriteTextRow(row, useSharedStrings: true);
                }
                writter1.CloseSheet();
            }

            using (BigExcelWritter writter2 = new BigExcelWritter(m2, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                writter2.CreateAndOpenSheet("s1");
                foreach (List<string> row in strings)
                {
                    writter2.WriteTextRow(row, useSharedStrings: false);
                }
                writter2.CloseSheet();
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


        #region private
        private static IEnumerable<Row> GetRows(Worksheet worksheet)
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

        private static IEnumerable<Cell> GetCells(Row row)
        {
            return row.ChildElements.OfType<Cell>();
        }

        private static string GetCellRealValue(Cell cell, WorkbookPart workbookPart)
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

        private static BigExcelWritter GetWritterStream(out MemoryStream stream)
        {
            stream = new MemoryStream();
            return new BigExcelWritter(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        }
        #endregion
    }
}
