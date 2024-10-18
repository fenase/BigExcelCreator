// Copyright (c) 2022-2024, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: rownum Validator Autofilter stylesheet finalizer inline unhiding gridlines

using BigExcelCreator.CommentsManager;
using BigExcelCreator.Exceptions;
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
        /// When <see langword="true"/>, shows gridlines on screen (default).
        /// When <see langword="false"/>, hides gridlines on screen.
        /// </summary>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        public bool ShowGridLinesInCurrentSheet
        {
            get => sheetOpen ? _showGridLinesInCurrentSheet : throw new NoOpenSheetException("Cannot get Grid Lines configuration because there is no open sheet");

            set => _showGridLinesInCurrentSheet = sheetOpen ? value : throw new NoOpenSheetException("Cannot set Grid Lines configuration because there is no open sheet");
        }
        private bool _showGridLinesInCurrentSheet = _showGridLinesDefault;
        private const bool _showGridLinesDefault = true;

        /// <summary>
        /// When <see langword="true"/>, shows row and column headings (default).
        /// When <see langword="false"/>, hides row and column headings.
        /// </summary>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        public bool ShowRowAndColumnHeadingsInCurrentSheet
        {
            get => sheetOpen ? _showRowAndColumnHeadingsInCurrentSheet : throw new NoOpenSheetException("Cannot get Headings configuration because there is no open sheet");

            set => _showRowAndColumnHeadingsInCurrentSheet = sheetOpen ? value : throw new NoOpenSheetException("Cannot set Headings configuration because there is no open sheet");
        }
        private bool _showRowAndColumnHeadingsInCurrentSheet = _showRowAndColumnHeadingsDefault;
        private const bool _showRowAndColumnHeadingsDefault = true;

        /// <summary>
        /// When <see langword="true"/>, Prints gridlines.
        /// When <see langword="false"/>, Doesn't print gridlines (default).
        /// </summary>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        public bool PrintGridLinesInCurrentSheet
        {
            get => sheetOpen ? _printGridLinesInCurrentSheet : throw new NoOpenSheetException("Cannot get Grid Lines print configuration because there is no open sheet");
            set => _printGridLinesInCurrentSheet = sheetOpen ? value : throw new NoOpenSheetException("Cannot set Grid Lines print configuration because there is no open sheet");
        }
        private bool _printGridLinesInCurrentSheet = _printGridLinesDefault;
        private const bool _printGridLinesDefault = false;

        /// <summary>
        /// When <see langword="true"/>, Prints row and column headings.
        /// When <see langword="false"/>, Doesn't print row and column headings (default).
        /// </summary>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        public bool PrintRowAndColumnHeadingsInCurrentSheet
        {
            get => sheetOpen ? _printRowAndColumnHeadingsInCurrentSheet : throw new NoOpenSheetException("Cannot get Headings print configuration because there is no open sheet");
            set => _printRowAndColumnHeadingsInCurrentSheet = sheetOpen ? value : throw new NoOpenSheetException("Cannot set Headings print configuration because there is no open sheet");
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

        private readonly List<Sheet> sheets = [];

        private DataValidations sheetDataValidations;

        private OpenXmlWriter workSheetPartWriter;

        private readonly List<string> SharedStringsList = [];

        private WorksheetPart workSheetPart;

        private CommentManager commentManager;

        private AutoFilter SheetAutoFilter;

        private SharedStringTablePart SharedStringTablePart;

        private readonly List<ConditionalFormatting> conditionalFormattingList = [];

        private readonly List<CellRange> SheetMergedCells = [];

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
                WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                // add styles to sheet
                workbookStylesPart.Stylesheet = stylesheet;
                workbookStylesPart.Stylesheet.Save();
            }

            SharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();

            SkipCellWhenEmpty = skipCellWhenEmpty;
        }
        #endregion

        /// <summary>
        /// Creates a new sheet and prepares the writer to use it.
        /// </summary>
        /// <param name="name">Names the sheet</param>
        /// <exception cref="SheetAlreadyOpenException">When a sheet is already open</exception>
        public void CreateAndOpenSheet(string name) => CreateAndOpenSheet(name, null, SheetStateValues.Visible);

        /// <summary>
        /// Creates a new sheet and prepares the writer to use it.
        /// </summary>
        /// <param name="name">Names the sheet</param>
        /// <param name="sheetState">Sets sheet visibility. <c>SheetStateValues.Visible</c> to list the sheet. <c>SheetStateValues.Hidden</c> to hide it. <c>SheetStateValues.VeryHidden</c> to hide it and prevent unhiding from the GUI.</param>
        /// <exception cref="SheetAlreadyOpenException">When a sheet is already open</exception>
        public void CreateAndOpenSheet(string name, SheetStateValues sheetState) => CreateAndOpenSheet(name, null, sheetState);

        /// <summary>
        /// Creates a new sheet and prepares the writer to use it.
        /// </summary>
        /// <param name="name">Names the sheet</param>
        /// <param name="columns">Use this to set the columns' width</param>
        /// <exception cref="SheetAlreadyOpenException">When a sheet is already open</exception>
        public void CreateAndOpenSheet(string name, IList<Column> columns) => CreateAndOpenSheet(name, columns, SheetStateValues.Visible);

        /// <summary>
        /// Creates a new sheet and prepares the writer to use it.
        /// </summary>
        /// <param name="name">Names the sheet</param>
        /// <param name="columns">Use this to set the columns' width</param>
        /// <param name="sheetState">Sets sheet visibility. <c>SheetStateValues.Visible</c> to list the sheet. <c>SheetStateValues.Hidden</c> to hide it. <c>SheetStateValues.VeryHidden</c> to hide it and prevent unhiding from the GUI.</param>
        /// <exception cref="SheetAlreadyOpenException">When a sheet is already open</exception>
        public void CreateAndOpenSheet(string name, IList<Column> columns, SheetStateValues sheetState)
        {
            if (sheetOpen) { throw new SheetAlreadyOpenException("Cannot open a new sheet. Please close current sheet before opening a new one"); }

            workSheetPart = Document.WorkbookPart.AddNewPart<WorksheetPart>();
            workSheetPartWriter = OpenXmlWriter.Create(workSheetPart);
            currentSheetName = name;
            workSheetPartWriter.WriteStartElement(new Worksheet());

            if (columns?.Count > 0)
            {
                workSheetPartWriter.WriteStartElement(new Columns());
                int columnIndex = 1;
                foreach (Column column in columns)
                {
                    List<OpenXmlAttribute> columnAttributes =
                    [
                        new OpenXmlAttribute("min", null, columnIndex.ToString(CultureInfo.InvariantCulture)),
                        new OpenXmlAttribute("max", null, columnIndex.ToString(CultureInfo.InvariantCulture)),
                        new OpenXmlAttribute("width", null, (column.Width ?? 11).ToString()),
                        new OpenXmlAttribute("customWidth", null, (column.CustomWidth ?? true).ToString()),
                        new OpenXmlAttribute("hidden", null, (column.Hidden ?? false).ToString()),
                    ];

                    workSheetPartWriter.WriteStartElement(new Column(), columnAttributes);
                    workSheetPartWriter.WriteEndElement();
                    ++columnIndex;
                }
                workSheetPartWriter.WriteEndElement();
            }

            workSheetPartWriter.WriteStartElement(new SheetData());
            sheetOpen = true;
            currentSheetState = sheetState;

            SetSheetDefault();
        }

        /// <summary>
        /// Closes a sheet
        /// </summary>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        public void CloseSheet()
        {
            if (!sheetOpen) { throw new NoOpenSheetException("There is no sheet to close"); }

            // write the end SheetData element
            workSheetPartWriter.WriteEndElement();

            WriteFilters();

            WriteMergedCells();

            WriteConditionalFormatting();

            WriteValidations();

            WritePrintOptions();

            // write the end Worksheet element
            workSheetPartWriter.WriteEndElement();

            workSheetPartWriter.Close();
            workSheetPartWriter = null;

            commentManager?.SaveComments(workSheetPart);

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

            sheetOpen = false;
            workSheetPart = null;
            commentManager = null;
            lastRowWritten = 0;
        }

        /// <summary>
        /// Creates a new row
        /// </summary>
        /// <param name="rownum">Row index</param>
        /// <exception cref="NoOpenSheetException">If there is no open sheet</exception>
        /// <exception cref="RowAlreadyOpenException">If already inside a row</exception>
        /// <exception cref="OutOfOrderWritingException">If attempting to write rows out of order</exception>
        public void BeginRow(int rownum)
        {
            BeginRow(rownum, false);
        }

        /// <summary>
        /// Creates a new row
        /// </summary>
        /// <param name="rownum">Row index</param>
        /// <param name="hidden">Hides the row when <see langword="true"/></param>
        /// <exception cref="NoOpenSheetException">If there is no open sheet</exception>
        /// <exception cref="RowAlreadyOpenException">If already inside a row</exception>
        /// <exception cref="OutOfOrderWritingException">If attempting to write rows out of order</exception>
        public void BeginRow(int rownum, bool hidden)
        {
            if (!sheetOpen) { throw new NoOpenSheetException("There is no open sheet to write a row to"); }
            if (rowOpen) { throw new RowAlreadyOpenException("A row is already open. Use EndRow to close it."); }
            if (rownum <= lastRowWritten) { throw new OutOfOrderWritingException("Writing rows out of order is not allowed"); }

            lastRowWritten = rownum;
            //create a new list of attributes
            List<OpenXmlAttribute> attributes =
            [
                // add the row index attribute to the list
                new OpenXmlAttribute("r", null, lastRowWritten.ToString(CultureInfo.InvariantCulture)),

                // Hide row if requested
                new OpenXmlAttribute("hidden", null, hidden ? "1" : "0"),
            ];

            //write the row start element with the row index attribute
            workSheetPartWriter.WriteStartElement(new Row(), attributes);
            rowOpen = true;
        }

        /// <summary>
        /// Creates a new row
        /// </summary>
        /// <exception cref="NoOpenSheetException">If there is no open sheet</exception>
        /// <exception cref="RowAlreadyOpenException">If already inside a row</exception>
        public void BeginRow()
        {
            BeginRow(false);
        }

        /// <summary>
        /// Creates a new row
        /// </summary>
        /// <param name="hidden">Hides the row when <see langword="true"/></param>
        /// <exception cref="NoOpenSheetException">If there is no open sheet</exception>
        /// <exception cref="RowAlreadyOpenException">If already inside a row</exception>
        public void BeginRow(bool hidden)
        {
            BeginRow(lastRowWritten + 1, hidden);
        }

        /// <summary>
        /// Closes a row
        /// </summary>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        public void EndRow()
        {
            if (!rowOpen) { throw new NoOpenRowException("There is no row to close"); }

            // write the end row element
            workSheetPartWriter.WriteEndElement();
            maxColumnNum = Math.Max(columnNum - 1, maxColumnNum);
            columnNum = 1;
            rowOpen = false;
        }

        /// <summary>
        /// Writes a string to a cell
        /// </summary>
        /// <param name="text">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="useSharedStrings">Write the value to the shared strings table. This might help reduce the output file size when the same text is shared multiple times among sheets.</param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        public void WriteTextCell(string text, int format = 0, bool useSharedStrings = false)
        {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!rowOpen) { throw new NoOpenRowException("There is no active row"); }

            if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(text)))
            {
                List<OpenXmlAttribute> attributes;
                if (useSharedStrings)
                {
                    string ssPos = AddTextToSharedStringsTable(text).ToString(CultureInfo.InvariantCulture);
                    attributes =
                    [
                        new OpenXmlAttribute("t", null, "s"),
                        new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.twoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
                        //styles
                        new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                    ];
                    //write the cell start element with the type and reference attributes
                    workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                    //write the cell value
                    workSheetPartWriter.WriteElement(new CellValue(ssPos));
                }
                else
                {
                    //reset the list of attributes
                    attributes =
                    [
                        // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                        new OpenXmlAttribute("t", null, "str"),
                        //add the cell reference attribute
                        new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.twoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
                        //styles
                        new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                    ];
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

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(sbyte number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        public void WriteNumberCell(byte number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        public void WriteNumberCell(short number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(ushort number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        public void WriteNumberCell(int number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(uint number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        public void WriteNumberCell(long number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(ulong number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        public void WriteNumberCell(float number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        public void WriteNumberCell(double number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        /// <summary>
        /// Writes a numerical value to a cell
        /// </summary>
        /// <param name="number">value to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        public void WriteNumberCell(decimal number, int format = 0)
        {
            WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);
        }

        private void WriteNumberCellInternal(string number, int format = 0)
        {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!rowOpen) { throw new NoOpenRowException("There is no active row"); }

            //reset the list of attributes
            List<OpenXmlAttribute> attributes =
            [
                //add the cell reference attribute
                new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.twoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
                //styles
                new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
            ];

            //write the cell start element with the type and reference attributes
            workSheetPartWriter.WriteStartElement(new Cell(), attributes);
            //write the cell value
            workSheetPartWriter.WriteElement(new CellValue(number.ToString(CultureInfo.InvariantCulture)));

            // write the end cell element
            workSheetPartWriter.WriteEndElement();

            columnNum++;
        }

        /// <summary>
        /// Writes a formula to a cell
        /// </summary>
        /// <param name="formula">formula to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">When there is no open row</exception>
        public void WriteFormulaCell(string formula, int format = 0)
        {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!rowOpen) { throw new NoOpenRowException("There is no active row"); }

            if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(formula)))
            {
                //reset the list of attributes
                List<OpenXmlAttribute> attributes =
                [
                    //add the cell reference attribute
                    new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.twoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
                    //styles
                    new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                ];

                //write the cell start element with the type and reference attributes
                workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                //write the cell value
                workSheetPartWriter.WriteElement(new CellFormula(formula?.ToUpperInvariant()));

                // write the end cell element
                workSheetPartWriter.WriteEndElement();
            }
            columnNum++;
        }

        /// <summary>
        /// Writes an entire text row at once
        /// </summary>
        /// <param name="texts">List of values to be written</param>
        /// <param name="format">Format index inside stylesheet. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Hides the row when <see langword="true"/></param>
        /// <param name="useSharedStrings">Write the value to the shared strings table. This might help reduce the output file size when the same text is shared multiple times among sheets.</param>
        /// <exception cref="ArgumentNullException">When list is <see langword="null"/></exception>
        /// <exception cref="NoOpenSheetException">If there is no open sheet</exception>
        /// <exception cref="RowAlreadyOpenException">If already inside a row</exception>
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
        /// <exception cref="NoOpenSheetException">If there is no open sheet</exception>
        /// <exception cref="RowAlreadyOpenException">If already inside a row</exception>
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
        /// <exception cref="NoOpenSheetException">If there is no open sheet</exception>
        /// <exception cref="RowAlreadyOpenException">If already inside a row</exception>
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
        /// <exception cref="NoOpenSheetException">When no open sheet </exception>
        /// <exception cref="SheetAlreadyHasFilterException">When there is already a filter an <paramref name="overwrite"/> is set to <see langword="false"/></exception>
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
        /// <exception cref="NoOpenSheetException">When no open sheet </exception>
        /// <exception cref="SheetAlreadyHasFilterException">When there is already a filter an <paramref name="overwrite"/> is set to <see langword="false"/></exception>
        /// <exception cref="ArgumentOutOfRangeException">When range height is not exactly one row</exception>
        public void AddAutofilter(CellRange range, bool overwrite = false)
        {
            if (!sheetOpen) { throw new NoOpenSheetException("Filters need to be assigned to a sheet"); }
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(range);
#else
            if (range == null) { throw new ArgumentNullException(nameof(range)); }
#endif
            if ((!overwrite) && SheetAutoFilter != null) { throw new SheetAlreadyHasFilterException("There is already a filter in use in current sheet. Set overwrite to true to replace it"); }
            if (range.Height != 1) { throw new ArgumentOutOfRangeException(nameof(range), "Range height must be 1"); }
            SheetAutoFilter = new AutoFilter() { Reference = range.RangeStringNoSheetName };
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
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
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
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        public void AddListValidator(CellRange range,
                                     string formula,
                                     bool allowBlank = true,
                                     bool showInputMessage = true,
                                     bool showErrorMessage = true)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.List, DataValidationOperatorValues.Equal, allowBlank, showInputMessage, showErrorMessage);

            AppendNewDataValidation(dataValidation, formula);
        }

        /// <summary>
        /// Adds an integer (whole) number validator to a range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="firstOperand"></param>
        /// <param name="validationType"></param>
        /// <param name="allowBlank"></param>
        /// <param name="showInputMessage"></param>
        /// <param name="showErrorMessage"></param>
        /// <param name="secondOperand"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="NoOpenSheetException"></exception>
        public void AddIntegerValidator(string range,
                                        int firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        int? secondOperand = null)
        {
            AddIntegerValidator(new CellRange(range),
                                firstOperand,
                                validationType,
                                allowBlank,
                                showInputMessage,
                                showErrorMessage,
                                secondOperand);
        }

        /// <summary>
        /// Adds an integer (whole) number validator to a range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="firstOperand"></param>
        /// <param name="validationType"></param>
        /// <param name="allowBlank"></param>
        /// <param name="showInputMessage"></param>
        /// <param name="showErrorMessage"></param>
        /// <param name="secondOperand"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="NoOpenSheetException"></exception>
        public void AddIntegerValidator(CellRange range,
                                        int firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        int? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Whole, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds a decimal number validator to a range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="firstOperand"></param>
        /// <param name="validationType"></param>
        /// <param name="allowBlank"></param>
        /// <param name="showInputMessage"></param>
        /// <param name="showErrorMessage"></param>
        /// <param name="secondOperand"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="NoOpenSheetException"></exception>
        public void AddDecimalValidator(string range,
                                        decimal firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        decimal? secondOperand = null)
        {
            AddDecimalValidator(new CellRange(range),
                                firstOperand,
                                validationType,
                                allowBlank,
                                showInputMessage,
                                showErrorMessage,
                                secondOperand);
        }

        /// <summary>
        /// Adds a decimal number validator to a range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="firstOperand"></param>
        /// <param name="validationType"></param>
        /// <param name="allowBlank"></param>
        /// <param name="showInputMessage"></param>
        /// <param name="showErrorMessage"></param>
        /// <param name="secondOperand"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="NoOpenSheetException"></exception>
        public void AddDecimalValidator(CellRange range,
                                        decimal firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        decimal? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Decimal, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds a comment to a cell
        /// </summary>
        /// <param name="text">Comment text</param>
        /// <param name="reference">Commented cell</param>
        /// <param name="author">Comment Author</param>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="author"/> is null or an empty string OR <paramref name="reference"/> is not a single cell</exception>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        /// <exception cref="InvalidRangeException">When <paramref name="reference"/> is not a valid range</exception>
        public void Comment(string text, string reference, string author = "BigExcelCreator")
        {
            if (string.IsNullOrEmpty(author)) { throw new ArgumentOutOfRangeException(nameof(author)); }
            CellRange cellRange = new(reference);
            if (!cellRange.IsSingleCellRange) { throw new ArgumentOutOfRangeException(nameof(reference), $"{nameof(reference)} must be a single cell range"); }
            if (!sheetOpen) { throw new NoOpenSheetException("Comments need to be placed on a sheet"); }

            commentManager ??= new();
            commentManager.Add(new CommentReference()
            {
                Cell = cellRange.RangeStringNoSheetName,
                Text = text,
                Author = author,
            });
        }

        /// <summary>
        /// Adds conditional formatting based on a formula
        /// </summary>
        /// <param name="reference">Cell to apply format to</param>
        /// <param name="formula">Formula. Format will be applied when this formula evaluates to true</param>
        /// <param name="format">Index of differential format in stylesheet. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <exception cref="ArgumentNullException">When formula is <see langword="null"/> or empty string</exception>
        /// <exception cref="ArgumentOutOfRangeException">When format is less than 0</exception>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        /// <exception cref="InvalidRangeException">When <paramref name="reference"/> is not a valid range</exception>
        public void AddConditionalFormattingFormula(string reference, string formula, int format)
        {
            if (!sheetOpen) { throw new NoOpenSheetException(ConstantsAndTexts.ConditionalFormattingMustBeOnSheet); }

            CellRange cellRange = new(reference);
            if (formula.IsNullOrWhiteSpace()) { throw new ArgumentNullException(nameof(formula)); }
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif

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

            conditionalFormattingRule.Append(new Formula { Text = formula });

            conditionalFormatting.Append(conditionalFormattingRule);

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
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        /// <exception cref="InvalidRangeException">When <paramref name="reference"/> is not a valid range</exception>
        public void AddConditionalFormattingCellIs(string reference, ConditionalFormattingOperatorValues @operator, string value, int format, string value2 = null)
        {
            CellRange cellRange = new(reference);

#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (value.IsNullOrWhiteSpace()) { throw new ArgumentNullException(nameof(value)); }
            if (!sheetOpen) { throw new NoOpenSheetException(ConstantsAndTexts.ConditionalFormattingMustBeOnSheet); }
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

            conditionalFormattingRule.Append(new Formula { Text = value });
            if (!value2.IsNullOrWhiteSpace()) { conditionalFormattingRule.Append(new Formula { Text = value2 }); }

            conditionalFormatting.Append(conditionalFormattingRule);

            conditionalFormattingList.Add(conditionalFormatting);
        }

        /// <summary>
        /// Adds conditional formatting to duplicated values
        /// </summary>
        /// <param name="reference">Cell to apply format to</param>
        /// <param name="format">Index of differential format in stylesheet. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">When format is less than 0</exception>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        /// <exception cref="InvalidRangeException">When <paramref name="reference"/> is not a valid range</exception>
        public void AddConditionalFormattingDuplicatedValues(string reference, int format)
        {
            CellRange cellRange = new(reference);

#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!sheetOpen) { throw new NoOpenSheetException(ConstantsAndTexts.ConditionalFormattingMustBeOnSheet); }

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

            conditionalFormatting.Append(conditionalFormattingRule);

            conditionalFormattingList.Add(conditionalFormatting);
        }

        /// <summary>
        /// Merges cells
        /// </summary>
        /// <param name="range">Cells to merge</param>
        /// <exception cref="ArgumentNullException">When <paramref name="range"/> is <see langword="null"/></exception>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        /// <exception cref="OverlappingRangesException">When trying to merge already merged cells</exception>
        public void MergeCells(CellRange range)
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(range);
#else
            if (range == null) { throw new ArgumentNullException(nameof(range)); }
#endif
            if (!sheetOpen) { throw new NoOpenSheetException(ConstantsAndTexts.ConditionalFormattingMustBeOnSheet); }

            if (SheetMergedCells.Exists(range.RangeOverlaps))
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
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
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

                workSheetPartWriter?.Dispose();
                Document.Dispose();

                if (SavingTo == SavingTo.stream)
                {
                    _ = Stream.Seek(0, SeekOrigin.Begin);
                }
            }
            open = false;
        }

        #region IDisposable
        private bool disposed;

        /// <summary>
        /// Saves and closes the document.
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                CloseDocument();
                if (disposing)
                {
                    // called via myClass.Dispose(). 
                    // OK to use any private object references
                }
                // Release unmanaged resources.
                // Set large fields to null.                
                disposed = true;
            }
        }

        /// <summary>
        /// Saves and closes the document.
        /// </summary>
        public void Dispose() // Implement IDisposable
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// The finalizer
        /// </summary>
        ~BigExcelWriter()
        {
            Dispose(false);
        }
        #endregion

        #region private methods
        private DataValidation AddValidatorCommon(CellRange range,
                                        DataValidationValues dataValidationValue,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true)
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(range);
#else
            if (range == null) { throw new ArgumentNullException(nameof(range)); }
#endif
            if (!sheetOpen) { throw new NoOpenSheetException("Validators need to be placed on a sheet"); }

            sheetDataValidations ??= new DataValidations();
            DataValidation dataValidation = new()
            {
                Type = dataValidationValue,
                AllowBlank = allowBlank,
                Operator = validationType,
                ShowInputMessage = showInputMessage,
                ShowErrorMessage = showErrorMessage,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = range.RangeString },
            };

            return dataValidation;
        }

        private void AppendNewDataValidation(DataValidation dataValidation, string firstOperand, string secondOperand = null)
        {
            sheetDataValidations ??= new DataValidations();

            Formula1 formula1 = new() { Text = firstOperand };
            dataValidation.Append(formula1);

            if (dataValidation.Operator.Value.RequiresSecondOperand())
            {
                Formula2 formula2 = new() { Text = secondOperand };
                dataValidation.Append(formula2);
            }
            sheetDataValidations.Append(dataValidation);
            sheetDataValidations.Count = (sheetDataValidations.Count ?? 0) + 1;
        }

        private void WriteFilters()
        {
            if (SheetAutoFilter == null) { return; }

            workSheetPartWriter.WriteElement(SheetAutoFilter);

            SheetAutoFilter = null;
        }

        private void WriteValidations()
        {
            if (sheetDataValidations == null) { return; }

            workSheetPartWriter.WriteStartElement(sheetDataValidations);
            foreach (DataValidation dataValidation in sheetDataValidations.ChildElements.Cast<DataValidation>())
            {
                workSheetPartWriter.WriteStartElement(dataValidation);
                workSheetPartWriter.WriteElement(dataValidation.Formula1);
                if (dataValidation.Formula2 != null) { workSheetPartWriter.WriteElement(dataValidation.Formula2); }
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
        }

        private void WriteSheetsAndClosePart()
        {
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

                worksheet.SheetViews = new SheetViews(sheetView);
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
