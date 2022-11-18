// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.CommentsManager;
using BigExcelCreator.Extensions;
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
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
using System.Threading.Tasks;
#endif

[assembly: CLSCompliant(true)]
[assembly: InternalsVisibleTo("Test")]
[assembly: InternalsVisibleTo("Test35")]
[assembly: InternalsVisibleTo("Test48")]

namespace BigExcelCreator
{
    /// <summary>
    /// This class writes Excel files directly using OpenXML SAX.
    /// Useful when trying to write tens of thousands of rows.
    /// <see href="https://www.nuget.org/packages/BigExcelCreator/#readme-body-tab">NuGet</see>
    /// <seealso href="https://github.com/fenase/BigExcelCreator">Source</seealso>
    /// </summary>
    public class BigExcelWriter : IDisposable
    {
        #region props
        /// <summary>
        /// Created file will be saved to: ...
        /// <para>(null when not saving to file)</para>
        /// </summary>
        public string Path { get; }

        /// <summary>
        /// Created file will be saved to: ...
        /// <para>(null when not saving to Stream)</para>
        /// </summary>
        public Stream Stream { get; }

        /// <summary>
        /// Where am I saving the file to (file or stream)?
        /// </summary>
        private SavingTo SavingTo { get; }

        /// <summary>
        /// Document type
        /// <para>only <c>SpreadsheetDocumentType.Workbook</c> is tested</para>
        /// </summary>
        public SpreadsheetDocumentType SpreadsheetDocumentType { get; private set; }

        /// <summary>
        /// The main document
        /// </summary>
        public SpreadsheetDocument Document { get; }

        /// <summary>
        /// When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written.
        /// When <see langword="false"/>, writing an empty value to a cell does nothing.
        /// </summary>
        public bool SkipCellWhenEmpty { get; set; }

        /// <summary>
        /// When <see langword="true""/>, shows gridlines on screen (default).
        /// When <see langword="false""/>, hides gridlines on screen.
        /// </summary>
        public bool ShowGridLinesInCurrentSheet
        {
            get => sheetOpen ? _showGridLinesInCurrentSheet : throw new InvalidOperationException("There is no open sheet");

            set => _showGridLinesInCurrentSheet = sheetOpen ? value : throw new InvalidOperationException("There is no open sheet");
        }
        private bool _showGridLinesInCurrentSheet = _showGridLinesDefault;
        private const bool _showGridLinesDefault = true;

        /// <summary>
        /// When <see langword="true""/>, shows row and column headings (default).
        /// When <see langword="false""/>, hides row and column headings.
        /// </summary>
        public bool ShowRowAndColumnHeadingsInCurrentSheet
        {
            get => sheetOpen ? _showRowAndColumnHeadingsInCurrentSheet : throw new InvalidOperationException("There is no open sheet");

            set => _showRowAndColumnHeadingsInCurrentSheet = sheetOpen ? value : throw new InvalidOperationException("There is no open sheet");
        }
        private bool _showRowAndColumnHeadingsInCurrentSheet = _showRowAndColumnHeadingsDefault;
        private const bool _showRowAndColumnHeadingsDefault = true;

        /// <summary>
        /// When <see langword="true""/>, Prints gridlines.
        /// When <see langword="false""/>, Doesn't print gridlines (default).
        /// </summary>
        public bool PrintGridLinesInCurrentSheet
        {
            get => sheetOpen ? _printGridLinesInCurrentSheet : throw new InvalidOperationException("There is no open sheet");
            set => _printGridLinesInCurrentSheet = sheetOpen ? value : throw new InvalidOperationException("There is no open sheet");
        }
        private bool _printGridLinesInCurrentSheet = _printGridLinesDefault;
        private const bool _printGridLinesDefault = false;

        /// <summary>
        /// When <see langword="true""/>, Prints row and column headings.
        /// When <see langword="false""/>, Doesn't print row and column headings (default).
        /// </summary>
        public bool PrintRowAndColumnHeadingsInCurrentSheet
        {
            get => sheetOpen ? _printRowAndColumnHeadingsInCurrentSheet : throw new InvalidOperationException("There is no open sheet");
            set => _printRowAndColumnHeadingsInCurrentSheet = sheetOpen ? value : throw new InvalidOperationException("There is no open sheet");
        }
        private bool _printRowAndColumnHeadingsInCurrentSheet = _printRowAndColumnHeadingsDefault;
        private const bool _printRowAndColumnHeadingsDefault = false;

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

        private OpenXmlWriter workSheetPartWriter;

        private readonly List<string> SharedStringsList = new();

        private WorksheetPart workSheetPart;

        private CommentManager commentManager;

        private AutoFilter SheetAutofilter;

        private SharedStringTablePart SharedStringTablePart;

        private readonly List<ConditionalFormatting> conditionalFormattingList = new();

        private readonly List<CellRange> SheetMergedCells = new();

#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
        private readonly List<Task> DocumentTasks = new();
        private readonly List<Task> SheetTasks = new();
#endif
        #endregion

        #region ctor
        /// <summary>
        /// Creates a document into <paramref name="stream"/>
        /// </summary>
        /// <param name="stream">Where to store the document. <c>MemoryStream</c> is recommended</param>
        /// <param name="spreadsheetDocumentType">Document type. Only <c>SpreadsheetDocumentType.Workbook</c> is tested</param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType)
        : this(stream, spreadsheetDocumentType, false) { }

        /// <summary>
        /// Creates a document into <paramref name="stream"/>
        /// </summary>
        /// <param name="stream">Where to store the document. <c>MemoryStream</c> is recommended</param>
        /// <param name="spreadsheetDocumentType">Document type. Only <c>SpreadsheetDocumentType.Workbook</c> is tested</param>
        /// <param name="stylesheet">A Stylesheet for the document. See <see cref="Styles.StyleList.GetStylesheet()"/></param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
        : this(stream, spreadsheetDocumentType, false, stylesheet) { }

        /// <summary>
        /// Creates a document into <paramref name="stream"/>
        /// </summary>
        /// <param name="stream">Where to store the document. <c>MemoryStream</c> is recommended</param>
        /// <param name="spreadsheetDocumentType">Document type. Only <c>SpreadsheetDocumentType.Workbook</c> is tested</param>
        /// <param name="skipCellWhenEmpty">When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty)
            : this(stream, spreadsheetDocumentType, skipCellWhenEmpty, new Stylesheet()) { }

        /// <summary>
        /// Creates a document into a file located in <paramref name="path"/>
        /// </summary>
        /// <param name="path">Path where the document will be saved</param>
        /// <param name="spreadsheetDocumentType">Document type. Only <c>SpreadsheetDocumentType.Workbook</c> is tested</param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType)
        : this(path, spreadsheetDocumentType, false) { }

        /// <summary>
        /// Creates a document into a file located in <paramref name="path"/>
        /// </summary>
        /// <param name="path">Path where the document will be saved</param>
        /// <param name="spreadsheetDocumentType">Document type. Only <c>SpreadsheetDocumentType.Workbook</c> is tested</param>
        /// <param name="stylesheet">A Stylesheet for the document. See <see cref="Styles.StyleList.GetStylesheet()"/></param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
        : this(path, spreadsheetDocumentType, false, stylesheet) { }

        /// <summary>
        /// Creates a document into a file located in <paramref name="path"/>
        /// </summary>
        /// <param name="path">Path where the document will be saved</param>
        /// <param name="spreadsheetDocumentType">Document type. Only <c>SpreadsheetDocumentType.Workbook</c> is tested</param>
        /// <param name="skipCellWhenEmpty">When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty)
            : this(path, spreadsheetDocumentType, skipCellWhenEmpty, new Stylesheet()) { }

        /// <summary>
        /// Creates a document into a file located in <paramref name="path"/>
        /// </summary>
        /// <param name="path">Path where the document will be saved</param>
        /// <param name="spreadsheetDocumentType">Document type. Only <c>SpreadsheetDocumentType.Workbook</c> is tested</param>
        /// <param name="skipCellWhenEmpty">When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        /// <param name="stylesheet">A Stylesheet for the document. See <see cref="Styles.StyleList.GetStylesheet()"/></param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            Path = path;
            SavingTo = SavingTo.file;
            Document = SpreadsheetDocument.Create(Path, spreadsheetDocumentType);
            CtorHelper(spreadsheetDocumentType, skipCellWhenEmpty, stylesheet);
        }

        /// <summary>
        /// Creates a document into <paramref name="stream"/>
        /// </summary>
        /// <param name="stream">Where to store the document. <c>MemoryStream</c> is recommended</param>
        /// <param name="spreadsheetDocumentType">Document type. Only <c>SpreadsheetDocumentType.Workbook</c> is tested</param>
        /// <param name="skipCellWhenEmpty">When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        /// <param name="stylesheet">A Stylesheet for the document. See <see cref="Styles.StyleList.GetStylesheet()"/></param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
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

            SharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();

            SkipCellWhenEmpty = skipCellWhenEmpty;
        }
        #endregion

        /// <summary>
        /// Creates a new sheet and prepares the writer to use it.
        /// </summary>
        /// <param name="name">Names the sheet</param>
        /// <param name="columns">Use this to set the columns' width</param>
        /// <param name="sheetState">Sets sheet visibility. <c>SheetStateValues.Visible</c> to list the sheet. <c>SheetStateValues.Hidden</c> to hide it. <c>SheetStateValues.VeryHidden</c> to hide it and prevent unhiding from the GUI.</param>
        /// <exception cref="InvalidOperationException">When a sheet is already open</exception>
        public void CreateAndOpenSheet(string name, IList<Column> columns = null,
                                       SheetStateValues sheetState = SheetStateValues.Visible)
        {
            if (!sheetOpen)
            {
                workSheetPart = Document.WorkbookPart.AddNewPart<WorksheetPart>();
                workSheetPartWriter = OpenXmlWriter.Create(workSheetPart);
                currentSheetName = name;
                workSheetPartWriter.WriteStartElement(new Worksheet());

                if (columns?.Count > 0)
                {
                    workSheetPartWriter.WriteStartElement(new Columns());
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

                        workSheetPartWriter.WriteStartElement(new Column(), atributosColumna);
                        workSheetPartWriter.WriteEndElement();
                        ++indiceColumna;
                    }
                    workSheetPartWriter.WriteEndElement();
                }

                workSheetPartWriter.WriteStartElement(new SheetData());
                sheetOpen = true;
                currentSheetState = sheetState;

                SetSheetDefault();
            }
            else
            {
                throw new InvalidOperationException("Sheet is already open");
            }
        }

        /// <summary>
        /// Closes a sheet
        /// </summary>
        /// <exception cref="InvalidOperationException">When there is no open sheet</exception>
        public void CloseSheet()
        {
            if (sheetOpen)
            {
                // write the end SheetData element
                workSheetPartWriter.WriteEndElement();

                WriteValidations();

                WriteFilters();

                WriteConditionalFormatting();

                WriteMergedCells();

                WritePrintOptions();

                // write the end Worksheet element
                workSheetPartWriter.WriteEndElement();

                workSheetPartWriter.Close();
                workSheetPartWriter = null;

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


                WritePageConfig(workSheetPart.Worksheet);



#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
                Task.WaitAll(SheetTasks.ToArray());
                SheetTasks.Clear();
#endif
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

        /// <summary>
        /// Creates a new row
        /// </summary>
        /// <param name="rownum">Row index</param>
        /// <exception cref="InvalidOperationException">If there is no open sheet OR already inside a row OR attempting to write rows out of order. See exception message for more details</exception>
        public void BeginRow(int rownum)
        {
            BeginRow(rownum, false);
        }

        /// <summary>
        /// Creates a new row
        /// </summary>
        /// <param name="rownum">Row index</param>
        /// <param name="hidden">Hides the row when <see langword="true"/></param>
        /// <exception cref="InvalidOperationException">If there is no open sheet OR already inside a row OR attempting to write rows out of order. See exception message for more details</exception>
        public void BeginRow(int rownum, bool hidden)
        {
            if (!sheetOpen) { throw new InvalidOperationException("There is no open sheet"); }
            if (rowOpen) { throw new InvalidOperationException("A row is already open. Use EndRow to close it."); }
            if (rownum <= lastRowWritten) { throw new InvalidOperationException("Out of order row writing is not allowed"); }

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
            workSheetPartWriter.WriteStartElement(new Row(), attributes);
            rowOpen = true;
        }

        /// <summary>
        /// Creates a new row
        /// </summary>
        /// <exception cref="InvalidOperationException">If there is no open sheet OR already inside a row. See exception message for more details</exception>
        public void BeginRow()
        {
            BeginRow(false);
        }

        /// <summary>
        /// Creates a new row
        /// </summary>
        /// <param name="hidden">Hides the row when <see langword="true"/></param>
        /// <exception cref="InvalidOperationException">If there is no open sheet OR already inside a row. See exception message for more details</exception>
        public void BeginRow(bool hidden)
        {
            BeginRow(lastRowWritten + 1, hidden);
        }

        /// <summary>
        /// Closes a row
        /// </summary>
        /// <exception cref="InvalidOperationException">When there is no open row</exception>
        public void EndRow()
        {
            if (rowOpen)
            {
                // write the end row element
                workSheetPartWriter.WriteEndElement();
                maxColumnNum = Math.Max(columnNum - 1, maxColumnNum);
                columnNum = 1;
                rowOpen = false;
            }
            else
            {
                throw new InvalidOperationException("There is no open row");
            }
        }

        /// <summary>
        /// Writes a string to a cell
        /// </summary>
        /// <param name="text">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="useSharedStrings">Write the value to the sharedstrings table. This might help reduce the output filesize when the same text is shared multiple times among sheets.</param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="InvalidOperationException">When there is no open row</exception>
        public void WriteTextCell(string text, int format = 0, bool useSharedStrings = false)
        {
            if (format < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(format));
            }

            if (rowOpen)
            {
                if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(text)))
                {
                    List<OpenXmlAttribute> attributes;
                    if (useSharedStrings)
                    {
                        string ssPos = AddTextToSharedStringsTable(text).ToString(CultureInfo.InvariantCulture);
                        attributes = new()
                        {
                            new OpenXmlAttribute("t", null, "s"),
                            new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture,"{0}{1}", Helpers.GetColumnName(columnNum), lastRowWritten)),
                            //styles
                            new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                        };
                        //write the cell start element with the type and reference attributes
                        workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                        //write the cell value
                        workSheetPartWriter.WriteElement(new CellValue(ssPos));
                    }
                    else
                    {
                        //reset the list of attributes
                        attributes = new()
                        {
                            // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                            new OpenXmlAttribute("t", null, "str"),
                            //add the cell reference attribute
                            new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture,"{0}{1}", Helpers.GetColumnName(columnNum), lastRowWritten)),
                            //styles
                            new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                        };
                        //write the cell start element with the type and reference attributes
                        workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                        //write the cell value
                        workSheetPartWriter.WriteElement(new CellValue(text));
                    }

                    // write the end cell element
                    workSheetPartWriter.WriteEndElement();

                }
                columnNum++;
            }
            else
            {
                throw new InvalidOperationException("There is no active row");
            }
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="InvalidOperationException">When there is no open row</exception>
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
                    //add the cell reference attribute
                    new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture,"{0}{1}", Helpers.GetColumnName(columnNum), lastRowWritten)),
                    //styles
                    new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                };

                //write the cell start element with the type and reference attributes
                workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                //write the cell value
                workSheetPartWriter.WriteElement(new CellValue(number.ToString(CultureInfo.InvariantCulture)));

                // write the end cell element
                workSheetPartWriter.WriteEndElement();

                columnNum++;
            }
            else
            {
                throw new InvalidOperationException("There is no active row");
            }
        }

        /// <summary>
        /// Writes a formula to a cell
        /// </summary>
        /// <param name="formula">formula to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="InvalidOperationException">When there is no open row</exception>
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
                        //styles
                        new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                    };

                    //write the cell start element with the type and reference attributes
                    workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                    //write the cell value
                    workSheetPartWriter.WriteElement(new CellFormula(formula?.ToUpperInvariant()));

                    // write the end cell element
                    workSheetPartWriter.WriteEndElement();
                }
                columnNum++;
            }
            else
            {
                throw new InvalidOperationException("There is no active row");
            }
        }

        /// <summary>
        /// Writes an entire text row at once
        /// </summary>
        /// <param name="texts">List of values to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Hides the row when <see langword="true"/></param>
        /// <param name="useSharedStrings">Write the value to the sharedstrings table. This might help reduce the output filesize when the same text is shared multiple times among sheets.</param>
        /// <exception cref="ArgumentNullException">When list is <see langword="null"/></exception>
        /// <exception cref="InvalidOperationException">If there is no open sheet OR already inside a row OR there is no open row. See exception message for more details</exception>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        public void WriteTextRow(IEnumerable<string> texts, int format = 0, bool hidden = false, bool useSharedStrings = false)
        {
            BeginRow(hidden);
            foreach (string text in texts ?? throw new ArgumentNullException(nameof(texts)))
            {
                WriteTextCell(text, format, useSharedStrings);
            }
            EndRow();
        }

        /// <summary>
        /// Writes an entire numerical row at once
        /// </summary>
        /// <param name="numbers">Lists of values to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Hides the row when <see langword="true"/></param>
        /// <exception cref="ArgumentNullException">When list is <see langword="null"/></exception>
        /// <exception cref="InvalidOperationException">If there is no open sheet OR already inside a row OR there is no open row. See exception message for more details</exception>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        public void WriteNumberRow(IEnumerable<float> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (float number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes an entire formula row at once
        /// </summary>
        /// <param name="formulas">List of formulas to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Hides the row when <see langword="true"/></param>
        /// <exception cref="ArgumentNullException">When list is <see langword="null"/></exception>
        /// <exception cref="InvalidOperationException">If there is no open sheet OR already inside a row OR there is no open row. See exception message for more details</exception>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        public void WriteFormulaRow(IEnumerable<string> formulas, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (string text in formulas ?? throw new ArgumentNullException(nameof(formulas)))
            {
                WriteFormulaCell(text, format);
            }
            EndRow();
        }

        /// <summary>
        /// Adds autofilter. Only one filter per sheet is allowed.
        /// </summary>
        /// <param name="range">Where to add the filter (header cells)</param>
        /// <param name="overwrite">Replace active filter</param>
        /// <exception cref="ArgumentNullException">Null range</exception>
        /// <exception cref="InvalidOperationException">When no open sheet OR there is already a filter an <paramref name="overwrite"/> is set to <see langword="false"/></exception>
        /// <exception cref="ArgumentOutOfRangeException">When range height is not exactly one row</exception>
        /// <exception cref="InvalidRangeException">When <paramref name="range"/> is not a valid range</exception>
        public void AddAutofilter(string range, bool overwrite = false)
        {
            AddAutofilter(new CellRange(range), overwrite);
        }

        /// <summary>
        /// Adds autofilter. Only one filter per sheet is allowed.
        /// </summary>
        /// <param name="range">Where to add the filter (header cells)</param>
        /// <param name="overwrite">Replace active filter</param>
        /// <exception cref="ArgumentNullException">Null range</exception>
        /// <exception cref="InvalidOperationException">When no open sheet OR there is already a filter an <paramref name="overwrite"/> is set to <see langword="false"/></exception>
        /// <exception cref="ArgumentOutOfRangeException">When range height is not exactly one row</exception>
        public void AddAutofilter(CellRange range, bool overwrite = false)
        {
            if (!sheetOpen) { throw new InvalidOperationException("There is no open sheet"); }
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

        /// <summary>
        /// Adds a list validator to a range based on a formula
        /// </summary>
        /// <param name="range">Cells to validate</param>
        /// <param name="formula">Validation formula</param>
        /// <param name="allowBlank"></param>
        /// <param name="showInputMessage"></param>
        /// <param name="showErrorMessage"></param>
        /// <exception cref="ArgumentNullException">When <paramref name="range"/> is null</exception>
        /// <exception cref="InvalidOperationException">When there is no open sheet</exception>
        /// <exception cref="InvalidRangeException">When <paramref name="range"/> is not a valid range</exception>
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

        /// <summary>
        /// Adds a list validator to a range based on a formula
        /// </summary>
        /// <param name="range">Cells to validate</param>
        /// <param name="formula">Validation formula</param>
        /// <param name="allowBlank"></param>
        /// <param name="showInputMessage"></param>
        /// <param name="showErrorMessage"></param>
        /// <exception cref="ArgumentNullException">When <paramref name="range"/> is null</exception>
        /// <exception cref="InvalidOperationException">When there is no open sheet</exception>
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

                dataValidation.Append(new[] { formula1 });
                sheetDataValidations.Append(new[] { dataValidation });
                sheetDataValidations.Count = (sheetDataValidations.Count ?? 0) + 1;
            }
            else
            {
                throw new InvalidOperationException("There is no open sheet");
            }
        }

        /// <summary>
        /// Adds a comment to a cell
        /// </summary>
        /// <param name="text">Comment text</param>
        /// <param name="reference">Commented cell</param>
        /// <param name="author">Comment Author</param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="author"/> is null or an empty string OR <paramref name="reference"/> is not a single cell</exception>
        /// <exception cref="InvalidOperationException">When there is no open sheet</exception>
        /// <exception cref="InvalidRangeException">When <paramref name="reference"/> is not a valid range</exception>
        public void Comment(string text, string reference, string author = "BigExcelCreator")
        {
            if (string.IsNullOrEmpty(author)) { throw new ArgumentOutOfRangeException(nameof(author)); }
            CellRange cellRange = new(reference);
            if (!cellRange.IsSingleCellRange) { throw new ArgumentOutOfRangeException(nameof(reference), $"{nameof(reference)} must be a single cell range"); }
            if (sheetOpen)
            {
                commentManager ??= new();
                commentManager.Add(new CommentReference()
                {
                    Cell = cellRange.RangeStringNoSheetName,
                    Text = text,
                    Author = author,
                });

            }
            else
            {
                throw new InvalidOperationException("There is no open sheet");
            }
        }

        /// <summary>
        /// Adds conditional formatting based on a formula
        /// </summary>
        /// <param name="reference">Cell to apply format to</param>
        /// <param name="formula">Formula. Format will be applied when this formula evaluates to true</param>
        /// <param name="format">Index of differential format in stylesheet. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <exception cref="ArgumentNullException">When formula is <see langword="null"/> or empty string</exception>
        /// <exception cref="ArgumentOutOfRangeException">When format is less than 0</exception>
        /// <exception cref="InvalidOperationException">When there is no open sheet</exception>
        /// <exception cref="InvalidRangeException">When <paramref name="reference"/> is not a valid range</exception>
        public void AddConditionalFormattingFormula(string reference, string formula, int format)
        {
            if (!sheetOpen) { throw new InvalidOperationException("There is no open sheet"); }

            CellRange cellRange = new(reference);
            if (formula.IsNullOrWhiteSpace()) { throw new ArgumentNullException(nameof(formula)); }
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }

            ConditionalFormatting conditionalFormatting = new()
            {
                SequenceOfReferences = new(new List<StringValue> { cellRange.RangeStringNoSheetName }),
            };

            ConditionalFormattingRule conditionalFormattingRule = new()
            {
                Type = ConditionalFormatValues.Expression,
                FormatId = (uint)format,
                Priority = conditionalFormattingList.Count + 1,
            };

            conditionalFormattingRule.Append(new[] { new Formula { Text = formula } });

            conditionalFormatting.Append(new[] { conditionalFormattingRule });

            conditionalFormattingList.Add(conditionalFormatting);
        }

        /// <summary>
        /// Adds conditional formatting based on cell value
        /// </summary>
        /// <param name="reference">Cell to apply format to</param>
        /// <param name="operator"></param>
        /// <param name="value">Compare cell value to this</param>
        /// <param name="format">Index of differential format in stylesheet. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <param name="value2">When <paramref name="operator"/> requires 2 parameters, compare cell value to this as second parameter</param>
        /// <exception cref="ArgumentOutOfRangeException">When format is less than 0</exception>
        /// <exception cref="ArgumentNullException">When <paramref name="value"/> is <see langword="null"/> OR <paramref name="operator"/> requires 2 arguments and <paramref name="value2"/> is <see langword="null"/></exception>
        /// <exception cref="InvalidOperationException">When there is no open sheet</exception>
        /// <exception cref="InvalidRangeException">When <paramref name="reference"/> is not a valid range</exception>
        public void AddConditionalFormattingCellIs(string reference, ConditionalFormattingOperatorValues @operator, string value, int format, string value2 = null)
        {
            CellRange cellRange = new(reference);

            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
            if (value.IsNullOrWhiteSpace()) { throw new ArgumentNullException(nameof(value)); }
            if (!sheetOpen) { throw new InvalidOperationException("There is no open sheet"); }
            if (new[] { ConditionalFormattingOperatorValues.Between, ConditionalFormattingOperatorValues.NotBetween }.Contains(@operator)
                && value2.IsNullOrWhiteSpace())
            {
                throw new ArgumentNullException(nameof(value2));
            }

            ConditionalFormatting conditionalFormatting = new()
            {
                SequenceOfReferences = new(new List<StringValue> { cellRange.RangeStringNoSheetName }),
            };

            ConditionalFormattingRule conditionalFormattingRule = new()
            {
                Type = ConditionalFormatValues.CellIs,
                @Operator = @operator,
                FormatId = (uint)format,
                Priority = conditionalFormattingList.Count + 1,
            };

            conditionalFormattingRule.Append(new[] { new Formula { Text = value } });
            if (!value2.IsNullOrWhiteSpace()) { conditionalFormattingRule.Append(new[] { new Formula { Text = value2 } }); }

            conditionalFormatting.Append(new[] { conditionalFormattingRule });

            conditionalFormattingList.Add(conditionalFormatting);
        }

        /// <summary>
        /// Adds conditional formatting to duplicated values
        /// </summary>
        /// <param name="reference">Cell to apply format to</param>
        /// <param name="format">Index of differential format in stylesheet. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When format is less than 0</exception>
        /// <exception cref="InvalidOperationException">When there is no open sheet</exception>
        /// <exception cref="InvalidRangeException">When <paramref name="reference"/> is not a valid range</exception>
        public void AddConditionalFormattingDuplicatedValues(string reference, int format)
        {
            CellRange cellRange = new(reference);

            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
            if (!sheetOpen) { throw new InvalidOperationException("There is no open sheet"); }

            ConditionalFormatting conditionalFormatting = new()
            {
                SequenceOfReferences = new(new List<StringValue> { cellRange.RangeStringNoSheetName }),
            };

            ConditionalFormattingRule conditionalFormattingRule = new()
            {
                Type = ConditionalFormatValues.DuplicateValues,
                FormatId = (uint)format,
                Priority = conditionalFormattingList.Count + 1,
            };

            conditionalFormatting.Append(new[] { conditionalFormattingRule });

            conditionalFormattingList.Add(conditionalFormatting);
        }

        /// <summary>
        /// Merges cells
        /// </summary>
        /// <param name="range">Cells to merge</param>
        /// <exception cref="ArgumentNullException">When <paramref name="range"/> is <see langword="null"/></exception>
        /// <exception cref="InvalidOperationException">When there is no open sheet</exception>
        /// <exception cref="OverlappingRangesException">When trying to merge already merged cells</exception>
        public void MergeCells(CellRange range)
        {
            if (range == null) { throw new ArgumentNullException(nameof(range)); }
            if (!sheetOpen) { throw new InvalidOperationException("There is no open sheet"); }

            if (SheetMergedCells.Any(range.RangeOverlaps))
            {
                throw new OverlappingRangesException();
            }
            else
            {
                SheetMergedCells.Add(range);
            }
        }

        /// <summary>
        /// Merges cells
        /// </summary>
        /// <param name="range">Cells to merge</param>
        /// <exception cref="InvalidRangeException">When <paramref name="range"/> is not a valid range</exception>
        /// <exception cref="InvalidOperationException">When there is no open sheet</exception>
        /// <exception cref="OverlappingRangesException">When trying to merge already merged cells</exception>
        public void MergeCells(string range)
        {
            MergeCells(new CellRange(range));
        }

        /// <summary>
        /// Closes the document
        /// </summary>
        public void CloseDocument()
        {
            if (open)
            {
                if (rowOpen) { EndRow(); }
                if (sheetOpen) { CloseSheet(); }

                WriteSharedStringsPart();
                WriteSheetsAndClosePart();

#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
                Task.WaitAll(SheetTasks.ToArray());
                SheetTasks.Clear();
                Task.WaitAll(DocumentTasks.ToArray());
                DocumentTasks.Clear();
#endif

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
                    workSheetPartWriter?.Dispose();
                    Document.Dispose();
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

        ~BigExcelWriter() // the finalizer
        {
            Dispose(false);
        }
        #endregion

        #region private methods
        private void WriteFilters()
        {
            if (SheetAutofilter == null) { return; }

            workSheetPartWriter.WriteElement(SheetAutofilter);

            SheetAutofilter = null;
        }



        private void WriteValidations()
        {
            if (sheetDataValidations == null) { return; }

            workSheetPartWriter.WriteStartElement(sheetDataValidations);
            foreach (DataValidation dataValidation in sheetDataValidations.ChildElements.Cast<DataValidation>())
            {
                workSheetPartWriter.WriteStartElement(dataValidation);
                workSheetPartWriter.WriteElement(dataValidation.Formula1);
                workSheetPartWriter.WriteEndElement();
            }
            workSheetPartWriter.WriteEndElement();

            sheetDataValidations = null;
        }

        private void WriteConditionalFormatting()
        {
            if (conditionalFormattingList == null || conditionalFormattingList.Count == 0) { return; }

            foreach (ConditionalFormatting conditionalFormatting in conditionalFormattingList)
            {
                workSheetPartWriter.WriteStartElement(conditionalFormatting);
                foreach (ConditionalFormattingRule conditionalFormattingRule in conditionalFormatting.ChildElements.OfType<ConditionalFormattingRule>())
                {
                    workSheetPartWriter.WriteStartElement(conditionalFormattingRule);
                    foreach (var item in conditionalFormattingRule.ChildElements)
                    {
                        workSheetPartWriter.WriteElement(item);
                    }
                    workSheetPartWriter.WriteEndElement();
                }
                workSheetPartWriter.WriteEndElement();
            }

            conditionalFormattingList.Clear();
        }

        private void WriteMergedCells()
        {
            if (SheetMergedCells == null || SheetMergedCells.Count == 0) { return; }

            workSheetPartWriter.WriteStartElement(new MergeCells());

            foreach (CellRange range in SheetMergedCells)
            {
                workSheetPartWriter.WriteElement(new MergeCell { Reference = range.RangeStringNoSheetName });
            }

            workSheetPartWriter.WriteEndElement();

            SheetMergedCells.Clear();
        }

        private int AddTextToSharedStringsTable(string text)
        {
            int pos = SharedStringsList.IndexOf(text);
            if (pos < 0)
            {
                pos = SharedStringsList.Count;
                SharedStringsList.Add(text);
            }
            return pos;
        }

        private void WriteSharedStringsPart()
        {
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
            Task task = new Task(() =>
            {
#endif
            using OpenXmlWriter SharedStringsWriter = OpenXmlWriter.Create(SharedStringTablePart);
            SharedStringsWriter.WriteStartElement(new SharedStringTable());
            foreach (string item in SharedStringsList)
            {
                SharedStringsWriter.WriteStartElement(new SharedStringItem());
                SharedStringsWriter.WriteElement(new Text(item));
                SharedStringsWriter.WriteEndElement();
            }
            SharedStringsWriter.WriteEndElement();
            SharedStringsWriter.Close();
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
            });
            task.Start();
            DocumentTasks.Add(task);
#endif
        }

        private void WriteSheetsAndClosePart()
        {
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
            Task task = new Task(() =>
            {
#endif
            using OpenXmlWriter workbookPartWriter = OpenXmlWriter.Create(Document.WorkbookPart);
            workbookPartWriter.WriteStartElement(new Workbook());
            workbookPartWriter.WriteStartElement(new Sheets());

            foreach (Sheet sheet in sheets)
            {
                workbookPartWriter.WriteElement(sheet);
            }

            // End Sheets
            workbookPartWriter.WriteEndElement();
            // End Workbook
            workbookPartWriter.WriteEndElement();
            workbookPartWriter.Close();
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
            });
            task.Start();
            DocumentTasks.Add(task);
#endif
        }

        private void SetSheetDefault()
        {
            _showGridLinesInCurrentSheet = _showGridLinesDefault;
            _showRowAndColumnHeadingsInCurrentSheet = _showRowAndColumnHeadingsDefault;
            _printRowAndColumnHeadingsInCurrentSheet = _printRowAndColumnHeadingsDefault;
            _printGridLinesInCurrentSheet = _printGridLinesDefault;
        }

        private void WritePageConfig(Worksheet worksheet)
        {
            if (_showGridLinesInCurrentSheet != _showGridLinesDefault || _showRowAndColumnHeadingsInCurrentSheet != _showRowAndColumnHeadingsDefault)
            {
                SheetView sheetView = new();
                if (_showGridLinesInCurrentSheet != _showGridLinesDefault) { sheetView.ShowGridLines = _showGridLinesInCurrentSheet; }
                if (_showRowAndColumnHeadingsInCurrentSheet != _showRowAndColumnHeadingsDefault) { sheetView.ShowRowColHeaders = _showRowAndColumnHeadingsInCurrentSheet; }
                sheetView.WorkbookViewId = 0;

                worksheet.SheetViews = new SheetViews(new[] { sheetView });
            }

        }

        private void WritePrintOptions()
        {
            if (_printGridLinesInCurrentSheet != _printGridLinesDefault || _printRowAndColumnHeadingsInCurrentSheet != _printRowAndColumnHeadingsDefault)
            {
                PrintOptions printOptions = new();
                if (_printGridLinesInCurrentSheet != _printGridLinesDefault) { printOptions.GridLines = _printGridLinesInCurrentSheet; }
                if (_printRowAndColumnHeadingsInCurrentSheet != _printRowAndColumnHeadingsDefault) { printOptions.Headings = _printRowAndColumnHeadingsInCurrentSheet; }
                workSheetPartWriter.WriteElement(printOptions);
            }
        }
        #endregion
    }

    internal enum SavingTo
    {
        file,
        stream,
    }
}
