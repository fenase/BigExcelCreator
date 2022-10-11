using BigExcelCreator.CommentsManager;
using BigExcelCreator.Ranges;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

[assembly: CLSCompliant(true)]
[assembly: InternalsVisibleTo("Test")]
[assembly: InternalsVisibleTo("Test35")]

namespace BigExcelCreator
{
    /// <summary>
    /// This class writes Excel files directly using OpenXML SAX.
    /// Useful when trying to write tens of thousands of rows.
    /// <see cref="https://www.nuget.org/packages/BigExcelCreator/#readme-body-tab">NuGet</see>
    /// <seealso cref="https://dev.azure.com/fenase/BigExcelCreator/_git/BigExcelCreator">Source</seealso>
    /// </summary>
    public class BigExcelWritter : IDisposable
    {
        #region props
        public string Path { get; }
        public Stream Stream { get; }
        private SavingTo SavingTo { get; }

        public SpreadsheetDocumentType SpreadsheetDocumentType { get; private set; }

        public SpreadsheetDocument Document { get; }

        public bool SkipCellWhenEmpty { get; set; }

        private bool sheetOpen;
        private string currentSheetName = "";
        private uint currentSheetId = 1;
        private SheetStateValues currentSheetState = SheetStateValues.Visible;
        private bool open = true;
        private int lastRowWritten;
        private bool rowOpen;
        private int columnNum = 1;
        private int maxColumnNum = 1;

        private readonly List<Sheet> sheets = new();

        private DataValidations sheetDataValidations;

        private OpenXmlWriter writer;

        private WorksheetPart workSheetPart;

        private CommentManager commentManager;

        private AutoFilter SheetAutofilter;
        #endregion

        #region ctor
        public BigExcelWritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType)
        : this(stream, spreadsheetDocumentType, false) { }

        public BigExcelWritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
        : this(stream, spreadsheetDocumentType, false, stylesheet) { }

        public BigExcelWritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty)
            : this(stream, spreadsheetDocumentType, skipCellWhenEmpty, new Stylesheet()) { }

        public BigExcelWritter(string path, SpreadsheetDocumentType spreadsheetDocumentType)
        : this(path, spreadsheetDocumentType, false) { }

        public BigExcelWritter(string path, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
        : this(path, spreadsheetDocumentType, false, stylesheet) { }

        public BigExcelWritter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty)
            : this(path, spreadsheetDocumentType, skipCellWhenEmpty, new Stylesheet()) { }

        public BigExcelWritter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            Path = path;
            SavingTo = SavingTo.file;
            Document = SpreadsheetDocument.Create(Path, spreadsheetDocumentType);
            CtorHelper(spreadsheetDocumentType, skipCellWhenEmpty, stylesheet);
        }

        public BigExcelWritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            Stream = stream;
            SavingTo = SavingTo.stream;
            Document = SpreadsheetDocument.Create(Stream, spreadsheetDocumentType);
            CtorHelper(spreadsheetDocumentType, skipCellWhenEmpty, stylesheet);
        }

        private void CtorHelper(SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            SpreadsheetDocumentType = spreadsheetDocumentType;
            WorkbookPart workbookPart = Document.AddWorkbookPart();

            if (workbookPart.WorkbookStylesPart == null)
            {
                WorkbookStylesPart wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();
                // add styles to sheet
                wbsp.Stylesheet = stylesheet;
                wbsp.Stylesheet.Save();
            }

            SkipCellWhenEmpty = skipCellWhenEmpty;
        }
        #endregion

        public void CreateAndOpenSheet(string name, IList<Column> columns = null,
                                       SheetStateValues sheetState = SheetStateValues.Visible)
        {
            if (!sheetOpen)
            {
                workSheetPart = Document.WorkbookPart.AddNewPart<WorksheetPart>();
                writer = OpenXmlWriter.Create(workSheetPart);
                currentSheetName = name;
                writer.WriteStartElement(new Worksheet());

                if (columns?.Count > 0)
                {
                    writer.WriteStartElement(new Columns());
                    int indiceColumna = 1;
                    foreach (Column column in columns)
                    {
                        List<OpenXmlAttribute> atributosColumna = new()
                        {
                            new OpenXmlAttribute("min", null, indiceColumna.ToString(CultureInfo.InvariantCulture)),
                            new OpenXmlAttribute("max", null, indiceColumna.ToString(CultureInfo.InvariantCulture)),
                            new OpenXmlAttribute("width", null, (column.Width ?? 11).ToString()),
                            new OpenXmlAttribute("customWidth", null, (column.CustomWidth ?? true).ToString()),
                            new OpenXmlAttribute("hidden", null, (column.Hidden ?? false).ToString()),
                        };

                        writer.WriteStartElement(new Column(), atributosColumna);
                        writer.WriteEndElement();
                        ++indiceColumna;
                    }
                    writer.WriteEndElement();
                }

                writer.WriteStartElement(new SheetData());
                sheetOpen = true;
                currentSheetState = sheetState;
            }
            else
            {
                throw new InvalidOperationException("Sheet is already open");
            }
        }

        public void CloseSheet()
        {
            if (sheetOpen)
            {
                // write the end SheetData element
                writer.WriteEndElement();
                // write validations
                WriteValidations();

                WriteFilters();

                // write the end Worksheet element
                writer.WriteEndElement();

                writer.Close();

                if (commentManager != null)
                {
                    commentManager.SaveComments(workSheetPart);
                }

                sheets.Add(new Sheet()
                {
                    Name = currentSheetName,
                    SheetId = currentSheetId++,
                    Id = Document.WorkbookPart.GetIdOfPart(workSheetPart),
                    State = currentSheetState,
                });

                currentSheetName = "";
                workSheetPart.Worksheet.SheetDimension = new SheetDimension() { Reference = $"A1:{Helpers.GetColumnName(maxColumnNum)}{Math.Max(1, lastRowWritten)}" };
                sheetOpen = false;
                workSheetPart = null;
                commentManager = null;
                lastRowWritten = 0;
            }
            else
            {
                throw new InvalidOperationException("There is no open sheet");
            }
        }


        public void BeginRow(int rownum)
        {
            BeginRow(rownum, false);
        }

        public void BeginRow(int rownum, bool hidden)
        {
            if (sheetOpen && !rowOpen)
            {
                if (rownum > lastRowWritten)
                {
                    lastRowWritten = rownum;
                    //create a new list of attributes
                    List<OpenXmlAttribute> attributes = new()
                    {
                        // add the row index attribute to the list
                        new OpenXmlAttribute("r", null, lastRowWritten.ToString(CultureInfo.InvariantCulture)),
                        
                        // Hide row if requested
                        new OpenXmlAttribute("hidden", null, hidden ? "1" : "0"),
                    };

                    //write the row start element with the row index attribute
                    writer.WriteStartElement(new Row(), attributes);
                    rowOpen = true;
                }
                else
                {
                    throw new InvalidOperationException("Out of order row writing is not allowed");
                }
            }
            else
            {
                throw new InvalidOperationException("A row is already open. Use EndRow to close it.");
            }
        }

        public void BeginRow()
        {
            BeginRow(false);
        }

        public void BeginRow(bool hidden)
        {
            BeginRow(lastRowWritten + 1, hidden);
        }

        public void EndRow()
        {
            if (rowOpen)
            {
                // write the end row element
                writer.WriteEndElement();
                maxColumnNum = Math.Max(columnNum - 1, maxColumnNum);
                columnNum = 1;
                rowOpen = false;
            }
            else
            {
                throw new InvalidOperationException("There is no row open");
            }
        }

        public void WriteTextCell(string text, int format = 0)
        {
            if (format < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(format));
            }

            if (rowOpen)
            {
                if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(text)))
                {
                    //reset the list of attributes
                    List<OpenXmlAttribute> attributes = new()
                {
                    // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                    new OpenXmlAttribute("t", null, "str"),
                    //add the cell reference attribute
                    new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture,"{0}{1}", Helpers.GetColumnName(columnNum), lastRowWritten)),
                    //estilos
                    new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                };

                    //write the cell start element with the type and reference attributes
                    writer.WriteStartElement(new Cell(), attributes);
                    //write the cell value
                    writer.WriteElement(new CellValue(text));

                    // write the end cell element
                    writer.WriteEndElement();
                }
                columnNum++;
            }
            else
            {
                throw new InvalidOperationException("There is no active row");
            }
        }

        public void WriteNumberCell(float number, int format = 0)
        {
            if (format < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(format));
            }

            if (rowOpen)
            {
                //reset the list of attributes
                List<OpenXmlAttribute> attributes = new()
            {
                //estilos
                new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
            };

                //write the cell start element with the type and reference attributes
                writer.WriteStartElement(new Cell(), attributes);
                //write the cell value
                writer.WriteElement(new CellValue(number));

                // write the end cell element
                writer.WriteEndElement();

                columnNum++;
            }
            else
            {
                throw new InvalidOperationException("There is no active row");
            }
        }

        public void WriteFormulaCell(string formula, int format = 0)
        {
            if (format < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(format));
            }

            if (rowOpen)
            {
                if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(formula)))
                {
                    //reset the list of attributes
                    List<OpenXmlAttribute> attributes = new()
                {
                    //add the cell reference attribute
                    new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture,"{0}{1}", Helpers.GetColumnName(columnNum), lastRowWritten)),
                    //estilos
                    new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                };

                    //write the cell start element with the type and reference attributes
                    writer.WriteStartElement(new Cell(), attributes);
                    //write the cell value
                    writer.WriteElement(new CellFormula(formula?.ToUpperInvariant()));

                    // write the end cell element
                    writer.WriteEndElement();
                }
                columnNum++;
            }
            else
            {
                throw new InvalidOperationException("There is no active row");
            }
        }

        public void WriteTextRow(IEnumerable<string> texts, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (string text in texts ?? throw new ArgumentNullException(nameof(texts)))
            {
                WriteTextCell(text, format);
            }
            EndRow();
        }

        public void WriteNumberRow(IEnumerable<float> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (float number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        public void WriteFormulaRow(IEnumerable<string> formulas, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (string text in formulas ?? throw new ArgumentNullException(nameof(formulas)))
            {
                WriteFormulaCell(text, format);
            }
            EndRow();
        }


        public void AddAutofilter(string range, bool overwrite = false)
        {
            AddAutofilter(new CellRange(range), overwrite);
        }

        public void AddAutofilter(CellRange range, bool overwrite = false)
        {
            if (range == null) { throw new ArgumentNullException(nameof(range)); }
            if ((!overwrite) && SheetAutofilter != null) { throw new InvalidOperationException("There is already a filter in use. Set owerwrite to true to replace it"); }
            if (range.Height != 1) { throw new ArgumentOutOfRangeException(nameof(range), "Range height must be 1"); }
            SheetAutofilter = new AutoFilter() { Reference = range.RangeStringNoSheetName };
        }


        [Obsolete("\"Please use AddListValidator instead.\"", true)]
        public void AddValidator(string range, string formula)
        {
            AddListValidator(range, formula);
        }

        public void AddListValidator(string range,
                                 string formula,
                                 bool allowBlank = true,
                                 bool showInputMessage = true,
                                 bool showErrorMessage = true)
        {
            AddListValidator(new CellRange(range),
                             formula,
                             allowBlank,
                             showInputMessage,
                             showErrorMessage);
        }

        public void AddListValidator(CellRange range,
                             string formula,
                             bool allowBlank = true,
                             bool showInputMessage = true,
                             bool showErrorMessage = true)
        {
            if (range == null) { throw new ArgumentNullException(nameof(range)); }
            if (sheetOpen)
            {
                sheetDataValidations ??= new DataValidations();
                DataValidation dataValidation = new()
                {
                    Type = DataValidationValues.List,
                    AllowBlank = allowBlank,
                    Operator = DataValidationOperatorValues.Equal,
                    ShowInputMessage = showInputMessage,
                    ShowErrorMessage = showErrorMessage,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range.RangeString },
                };

                Formula1 formula1 = new() { Text = formula };

                dataValidation.Append(formula1);
                sheetDataValidations.Append(dataValidation);
                sheetDataValidations.Count = (sheetDataValidations.Count ?? 0) + 1;
            }
            else
            {
                throw new InvalidOperationException("There is no open sheet");
            }
        }

        public void Comment(string text, string reference, string author = "BigExcelCreator")
        {
            if (string.IsNullOrEmpty(author)) { throw new ArgumentOutOfRangeException(nameof(author)); }
            if (sheetOpen)
            {
                commentManager ??= new();
                commentManager.Add(new CommentReference()
                {
                    Cell = reference,
                    Text = text,
                    Author = author,
                });

            }
            else
            {
                throw new InvalidOperationException("There is no open sheet");
            }
        }

        public void CloseDocument()
        {
            if (open)
            {
                writer = OpenXmlWriter.Create(Document.WorkbookPart);
                writer.WriteStartElement(new Workbook());
                writer.WriteStartElement(new Sheets());

                foreach (Sheet sheet in sheets)
                {
                    writer.WriteElement(sheet);
                }

                // End Sheets
                writer.WriteEndElement();
                // End Workbook
                writer.WriteEndElement();
                writer.Close();
                Document.Close();

                if (SavingTo == SavingTo.stream)
                {
                    _ = Stream.Seek(0, SeekOrigin.Begin);
                }
            }
            open = false;
        }

        #region IDisposable
        private bool disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                CloseDocument();
                if (disposing)
                {
                    // called via myClass.Dispose(). 
                    // OK to use any private object references
                    Document.Dispose();
                    writer.Dispose();
                }
                // Release unmanaged resources.
                // Set large fields to null.                
                disposed = true;
            }
        }

        public void Dispose() // Implement IDisposable
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~BigExcelWritter() // the finalizer
        {
            Dispose(false);
        }
        #endregion

        private void WriteFilters()
        {
            if (SheetAutofilter == null) { return; }

            writer.WriteElement(SheetAutofilter);

            SheetAutofilter = null;
        }



        private void WriteValidations()
        {
            if (sheetDataValidations == null) { return; }

            writer.WriteStartElement(sheetDataValidations);
            foreach (DataValidation item in sheetDataValidations.ChildElements.Cast<DataValidation>())
            {
                writer.WriteStartElement(item);
                writer.WriteElement(item.Formula1);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();

            sheetDataValidations = null;
        }
    }

    internal enum SavingTo
    {
        file,
        stream,
    }
}
