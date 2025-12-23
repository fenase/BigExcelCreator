// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.ClassAttributes;
using BigExcelCreator.CommentsManager;
using BigExcelCreator.Enums;
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
using System.Reflection;

namespace BigExcelCreator
{
    /// <summary>
    /// This class writes Excel files directly using OpenXML SAX.
    /// Useful when trying to write tens of thousands of rows.
    /// </summary>
    /// <remarks>
    /// <para><see href="https://www.nuget.org/packages/BigExcelCreator">NuGet</see></para>
    /// <para><seealso href="https://github.com/fenase/BigExcelCreator">Source</seealso></para>
    /// <para><seealso href="https://fenase.github.io/BigExcelCreator/api/BigExcelCreator.BigExcelWriter.html">API</seealso></para>
    /// <para><seealso href="https://fenase.github.io/projects/BigExcelCreator">Site</seealso></para>
    /// </remarks>
    public class BigExcelWriter : IDisposable
    {
        #region props
        /// <summary>
        /// Gets the file path where the Excel document is being saved.
        /// <para>(null when not saving to file)</para>
        /// </summary>
        public string Path { get; }

        /// <summary>
        /// Gets the Stream where the Excel document is being saved.
        /// <para>(null when not saving to Stream)</para>
        /// </summary>
        public Stream Stream { get; }

        /// <summary>
        /// Where am I saving the file to (file or stream)?
        /// </summary>
        private SavingTo SavingTo { get; }

        /// <summary>
        /// Gets the type of the spreadsheet document (e.g., Workbook, Template).
        /// <para>only <c>SpreadsheetDocumentType.Workbook</c> is tested</para>
        /// </summary>
        public SpreadsheetDocumentType SpreadsheetDocumentType { get; private set; }

        /// <summary>
        /// Gets the SpreadsheetDocument object representing the Excel document.
        /// </summary>
        public SpreadsheetDocument Document { get; }

        /// <summary>
        /// Gets or sets a value indicating whether to skip cells when they are empty.
        /// </summary>
        /// <remarks>
        /// When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written.
        /// When <see langword="false"/>, writing an empty value to a cell does nothing.
        /// </remarks>
        public bool SkipCellWhenEmpty { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to show grid lines in the current sheet.
        /// </summary>
        /// <remarks>
        /// When <see langword="true"/>, shows gridlines on screen (default).
        /// When <see langword="false"/>, hides gridlines on screen.
        /// </remarks>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        public bool ShowGridLinesInCurrentSheet
        {
            get => sheetOpen ? _showGridLinesInCurrentSheet : throw new NoOpenSheetException("Cannot get Grid Lines configuration because there is no open sheet");

            set => _showGridLinesInCurrentSheet = sheetOpen ? value : throw new NoOpenSheetException("Cannot set Grid Lines configuration because there is no open sheet");
        }
        private bool _showGridLinesInCurrentSheet = _showGridLinesDefault;
        private const bool _showGridLinesDefault = true;

        /// <summary>
        /// Gets or sets a value indicating whether to show row and column headings in the current sheet.
        /// </summary>
        /// <remarks>
        /// When <see langword="true"/>, shows row and column headings (default).
        /// When <see langword="false"/>, hides row and column headings.
        /// </remarks>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        public bool ShowRowAndColumnHeadingsInCurrentSheet
        {
            get => sheetOpen ? _showRowAndColumnHeadingsInCurrentSheet : throw new NoOpenSheetException("Cannot get Headings configuration because there is no open sheet");

            set => _showRowAndColumnHeadingsInCurrentSheet = sheetOpen ? value : throw new NoOpenSheetException("Cannot set Headings configuration because there is no open sheet");
        }
        private bool _showRowAndColumnHeadingsInCurrentSheet = _showRowAndColumnHeadingsDefault;
        private const bool _showRowAndColumnHeadingsDefault = true;

        /// <summary>
        /// Gets or sets a value indicating whether to print grid lines in the current sheet.
        /// </summary>
        /// <remarks>
        /// When <see langword="true"/>, Prints gridlines.
        /// When <see langword="false"/>, Doesn't print gridlines (default).
        /// </remarks>
        /// <exception cref="NoOpenSheetException">When there is no open sheet</exception>
        public bool PrintGridLinesInCurrentSheet
        {
            get => sheetOpen ? _printGridLinesInCurrentSheet : throw new NoOpenSheetException("Cannot get Grid Lines print configuration because there is no open sheet");
            set => _printGridLinesInCurrentSheet = sheetOpen ? value : throw new NoOpenSheetException("Cannot set Grid Lines print configuration because there is no open sheet");
        }
        private bool _printGridLinesInCurrentSheet = _printGridLinesDefault;
        private const bool _printGridLinesDefault = false;

        /// <summary>
        /// Gets or sets a value indicating whether to print row and column headings in the current sheet.
        /// </summary>
        /// <remarks>
        /// When <see langword="true"/>, Prints row and column headings.
        /// When <see langword="false"/>, Doesn't print row and column headings (default).
        /// </remarks>
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

        private readonly HashSet<string> SheetNames = [];

        #endregion

        #region ctor
        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified stream and spreadsheet document type.
        /// </summary>
        /// <remarks>Initializes a new Workbook</remarks>
        /// <param name="stream">The stream to write the Excel document to.</param>
        public BigExcelWriter(Stream stream)
        : this(stream, SpreadsheetDocumentType.Workbook) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified stream and spreadsheet document type.
        /// </summary>
        /// <param name="stream">The stream to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType)
        : this(stream, spreadsheetDocumentType, false) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified stream, spreadsheet document type, and stylesheet.
        /// </summary>
        /// <remarks>Initializes a new Workbook</remarks>
        /// <param name="stream">The stream to write the Excel document to.</param>
        /// <param name="stylesheet">The stylesheet to apply to the Excel document. See <see cref="Styles.StyleList.GetStylesheet()"/>.</param>
        public BigExcelWriter(Stream stream, Stylesheet stylesheet)
                : this(stream, SpreadsheetDocumentType.Workbook, stylesheet) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified stream, spreadsheet document type, and stylesheet.
        /// </summary>
        /// <param name="stream">The stream to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="stylesheet">The stylesheet to apply to the Excel document. See <see cref="Styles.StyleList.GetStylesheet()"/>.</param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
                : this(stream, spreadsheetDocumentType, false, stylesheet) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified stream, spreadsheet document type, and a flag indicating whether to skip cells when they are empty.
        /// </summary>
        /// <remarks>Initializes a new Workbook</remarks>
        /// <param name="stream">The stream to write the Excel document to.</param>
        /// <param name="skipCellWhenEmpty">A flag indicating whether to skip cells when they are empty. When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        public BigExcelWriter(Stream stream, bool skipCellWhenEmpty)
            : this(stream, SpreadsheetDocumentType.Workbook, skipCellWhenEmpty) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified stream, spreadsheet document type, and a flag indicating whether to skip cells when they are empty.
        /// </summary>
        /// <param name="stream">The stream to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="skipCellWhenEmpty">A flag indicating whether to skip cells when they are empty. When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty)
            : this(stream, spreadsheetDocumentType, skipCellWhenEmpty, new Stylesheet()) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified file path and spreadsheet document type.
        /// </summary>
        /// <remarks>Initializes a new Workbook</remarks>
        /// <param name="path">The file path to write the Excel document to.</param>
        public BigExcelWriter(string path)
        : this(path, SpreadsheetDocumentType.Workbook) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified file path and spreadsheet document type.
        /// </summary>
        /// <param name="path">The file path to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType)
        : this(path, spreadsheetDocumentType, false) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified file path, spreadsheet document type, and stylesheet.
        /// <remarks>Initializes a new Workbook</remarks>
        /// </summary>
        /// <param name="path">The file path to write the Excel document to.</param>
        /// <param name="stylesheet">The stylesheet to apply to the Excel document. See <see cref="Styles.StyleList.GetStylesheet()"/></param>
        public BigExcelWriter(string path, Stylesheet stylesheet)
        : this(path, SpreadsheetDocumentType.Workbook, stylesheet) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified file path, spreadsheet document type, and stylesheet.
        /// </summary>
        /// <param name="path">The file path to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="stylesheet">The stylesheet to apply to the Excel document. See <see cref="Styles.StyleList.GetStylesheet()"/></param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
        : this(path, spreadsheetDocumentType, false, stylesheet) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified file path, spreadsheet document type, and a flag indicating whether to skip cells when they are empty.
        /// </summary>
        /// <remarks>Initializes a new Workbook</remarks>
        /// <param name="path">The file path to write the Excel document to.</param>
        /// <param name="skipCellWhenEmpty">A flag indicating whether to skip cells when they are empty. When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        public BigExcelWriter(string path, bool skipCellWhenEmpty)
            : this(path, SpreadsheetDocumentType.Workbook, skipCellWhenEmpty) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified file path, spreadsheet document type, and a flag indicating whether to skip cells when they are empty.
        /// </summary>
        /// <param name="path">The file path to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="skipCellWhenEmpty">A flag indicating whether to skip cells when they are empty. When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty)
            : this(path, spreadsheetDocumentType, skipCellWhenEmpty, new Stylesheet()) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified file path, spreadsheet document type, a flag indicating whether to skip cells when they are empty, and a stylesheet.
        /// </summary>
        /// <param name="path">The file path to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="skipCellWhenEmpty">A flag indicating whether to skip cells when they are empty. When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        /// <param name="stylesheet">The stylesheet to apply to the Excel document. See <see cref="Styles.StyleList.GetStylesheet()"/>.</param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            ThrowIfInvalidSpreadsheetDocumentType(spreadsheetDocumentType);
            Path = path;
            SavingTo = SavingTo.file;
            Document = SpreadsheetDocument.Create(Path, spreadsheetDocumentType);
            CtorHelper(spreadsheetDocumentType, skipCellWhenEmpty, stylesheet);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified stream, spreadsheet document type, a flag indicating whether to skip cells when they are empty, and a stylesheet.
        /// </summary>
        /// <param name="stream">The stream to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="skipCellWhenEmpty">A flag indicating whether to skip cells when they are empty. When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        /// <param name="stylesheet">The stylesheet to apply to the Excel document. See <see cref="Styles.StyleList.GetStylesheet()"/>.</param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            ThrowIfInvalidSpreadsheetDocumentType(spreadsheetDocumentType);
            Stream = stream;
            SavingTo = SavingTo.stream;
            Document = SpreadsheetDocument.Create(Stream, spreadsheetDocumentType);
            CtorHelper(spreadsheetDocumentType, skipCellWhenEmpty, stylesheet);
        }

        private void CtorHelper(SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            disposed = false; // reset the disposed flag
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
        /// Creates and opens a new sheet with the specified name, and prepares the writer to use it.
        /// </summary>
        /// <param name="name">The name of the sheet to create and open.</param>
        /// <exception cref="SheetAlreadyOpenException">Thrown when a sheet is already open and not closed before opening a new one.</exception>
        /// <exception cref="SheetNameCannotBeEmptyException">Thrown when the sheet name is null or empty.</exception>
        /// <exception cref="SheetWithSameNameAlreadyExistsException">Thrown when a sheet with the same name already exists.</exception>
        public void CreateAndOpenSheet(string name) => CreateAndOpenSheet(name, null, SheetStateValues.Visible);

        /// <summary>
        /// Creates and opens a new sheet with the specified name, and sheet state, and prepares the writer to use it.
        /// </summary>
        /// <param name="name">The name of the sheet to create and open.</param>
        /// <param name="sheetState">Sets sheet visibility. <c>SheetStateValues.Visible</c> to list the sheet. <c>SheetStateValues.Hidden</c> to hide it. <c>SheetStateValues.VeryHidden</c> to hide it and prevent unhiding from the GUI.</param>
        /// <exception cref="SheetAlreadyOpenException">Thrown when a sheet is already open and not closed before opening a new one.</exception>
        /// <exception cref="SheetNameCannotBeEmptyException">Thrown when the sheet name is null or empty.</exception>
        /// <exception cref="SheetWithSameNameAlreadyExistsException">Thrown when a sheet with the same name already exists.</exception>
        public void CreateAndOpenSheet(string name, SheetStateValues sheetState) => CreateAndOpenSheet(name, null, sheetState);

        /// <summary>
        /// Creates and opens a new sheet with the specified name and columns, and prepares the writer to use it.
        /// </summary>
        /// <param name="name">The name of the sheet to create and open.</param>
        /// <param name="columns">The columns to add to the sheet. Can be null. Use this to set the columns' width.</param>
        /// <exception cref="SheetAlreadyOpenException">Thrown when a sheet is already open and not closed before opening a new one.</exception>
        /// <exception cref="SheetNameCannotBeEmptyException">Thrown when the sheet name is null or empty.</exception>
        /// <exception cref="SheetWithSameNameAlreadyExistsException">Thrown when a sheet with the same name already exists.</exception>
        public void CreateAndOpenSheet(string name, IList<Column> columns) => CreateAndOpenSheet(name, columns, SheetStateValues.Visible);

        /// <summary>
        /// Creates and opens a new sheet with the specified name, columns, and sheet state, and prepares the writer to use it.
        /// </summary>
        /// <param name="name">The name of the sheet to create and open.</param>
        /// <param name="columns">The columns to add to the sheet. Can be null. Use this to set the columns' width.</param>
        /// <param name="sheetState">Sets sheet visibility. <c>SheetStateValues.Visible</c> to list the sheet. <c>SheetStateValues.Hidden</c> to hide it. <c>SheetStateValues.VeryHidden</c> to hide it and prevent unhiding from the GUI.</param>
        /// <exception cref="SheetAlreadyOpenException">Thrown when a sheet is already open and not closed before opening a new one.</exception>
        /// <exception cref="SheetNameCannotBeEmptyException">Thrown when the sheet name is null or empty.</exception>
        /// <exception cref="SheetWithSameNameAlreadyExistsException">Thrown when a sheet with the same name already exists.</exception>
        public void CreateAndOpenSheet(string name, IList<Column> columns, SheetStateValues sheetState)
        {
            if (sheetOpen) { throw new SheetAlreadyOpenException("Cannot open a new sheet. Please close current sheet before opening a new one"); }

            if (string.IsNullOrEmpty(name)) { throw new SheetNameCannotBeEmptyException("Sheet name cannot be null or empty"); }
            if (SheetNames.Contains(name, StringComparer.OrdinalIgnoreCase)) { throw new SheetWithSameNameAlreadyExistsException("A sheet with the same name already exists"); }
            _ = SheetNames.Add(name);

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

            SetSheetDefaults();
        }

        /// <summary>
        /// Closes the currently open sheet.
        /// </summary>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to close.</exception>
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
        /// Begins a new row in the currently open sheet.
        /// </summary>
        /// <param name="rownum">The row number to begin.</param>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="OutOfOrderWritingException">Thrown when writing rows out of order is attempted.</exception>
        public void BeginRow(int rownum) => BeginRow(rownum, false);

        /// <summary>
        /// Begins a new row in the currently open sheet.
        /// </summary>
        /// <param name="rownum">The row number to begin.</param>
        /// <param name="hidden">Indicates whether the row should be hidden.</param>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="OutOfOrderWritingException">Thrown when writing rows out of order is attempted.</exception>
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
        /// Begins a new row in the currently open sheet.
        /// </summary>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="OutOfOrderWritingException">Thrown when writing rows out of order is attempted.</exception>
        public void BeginRow() => BeginRow(false);

        /// <summary>
        /// Begins a new row in the currently open sheet.
        /// </summary>
        /// <param name="hidden">Indicates whether the row should be hidden.</param>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="OutOfOrderWritingException">Thrown when writing rows out of order is attempted.</exception>
        public void BeginRow(bool hidden) => BeginRow(lastRowWritten + 1, hidden);

        /// <summary>
        /// Ends the currently open row in the sheet.
        /// </summary>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to end.</exception>
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
        /// Writes a text cell to the currently open row in the sheet.
        /// </summary>
        /// <param name="text">The text to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <param name="useSharedStrings">Indicates whether to write the value to the shared strings table. This might help reduce the output file size when the same text is shared multiple times among sheets. Default is false.</param>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        public void WriteTextCell(string text, int format = 0, bool useSharedStrings = false)
        {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!rowOpen) { throw new NoOpenRowException(ConstantsAndTexts.NoActiveRow); }

            if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(text)))
            {
                List<OpenXmlAttribute> attributes;
                if (useSharedStrings)
                {
                    string ssPos = AddTextToSharedStringsTable(text).ToString(CultureInfo.InvariantCulture);
                    attributes =
                    [
                        new OpenXmlAttribute("t", null, "s"),
                        new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
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
                        new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
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
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(sbyte number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(byte number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(short number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(ushort number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(int number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(uint number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(long number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(ulong number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(float number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(double number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(decimal number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        private void WriteNumberCell(object number, int format = 0)
        {
            switch (number)
            {
                case sbyte v: WriteNumberCell(v, format); break;
                case byte v: WriteNumberCell(v, format); break;
                case short v: WriteNumberCell(v, format); break;
                case ushort v: WriteNumberCell(v, format); break;
                case int v: WriteNumberCell(v, format); break;
                case uint v: WriteNumberCell(v, format); break;
                case long v: WriteNumberCell(v, format); break;
                case ulong v: WriteNumberCell(v, format); break;
                case float v: WriteNumberCell(v, format); break;
                case double v: WriteNumberCell(v, format); break;
                case decimal v: WriteNumberCell(v, format); break;
                default:
                    throw new ArgumentException(
                        $"Unsupported numeric type '{number.GetType().FullName}'. Supported types include: sbyte, byte, short, ushort, int, uint, long, ulong, float, double, decimal.",
                        nameof(number));
            }
        }

        /// <summary>
        /// Writes a formula cell to the currently open row in the sheet.
        /// </summary>
        /// <param name="formula">The formula to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteFormulaCell(string formula, int format = 0)
        {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!rowOpen) { throw new NoOpenRowException(ConstantsAndTexts.NoActiveRow); }

            if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(formula)))
            {
                //reset the list of attributes
                List<OpenXmlAttribute> attributes =
                [
                    //add the cell reference attribute
                    new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
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
        /// Writes a row of text cells to the currently open sheet.
        /// </summary>
        /// <param name="texts">The collection of text strings to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <param name="useSharedStrings">Indicates whether to write the value to the shared strings table. This might help reduce the output file size when the same text is shared multiple times among sheets. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the texts collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
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
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        [CLSCompliant(false)]
        public void WriteNumberRow(IEnumerable<sbyte> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (sbyte number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        public void WriteNumberRow(IEnumerable<byte> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (byte number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        public void WriteNumberRow(IEnumerable<short> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (short number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        [CLSCompliant(false)]
        public void WriteNumberRow(IEnumerable<ushort> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (ushort number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        public void WriteNumberRow(IEnumerable<int> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (int number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        [CLSCompliant(false)]
        public void WriteNumberRow(IEnumerable<uint> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (uint number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        public void WriteNumberRow(IEnumerable<long> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (long number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        [CLSCompliant(false)]
        public void WriteNumberRow(IEnumerable<ulong> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (ulong number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
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
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        public void WriteNumberRow(IEnumerable<double> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (double number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of cells with numerical values to the currently open sheet.
        /// </summary>
        /// <param name="numbers">The collection of numbers to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the numbers collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        public void WriteNumberRow(IEnumerable<decimal> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (decimal number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of formula cells to the currently open sheet.
        /// </summary>
        /// <param name="formulas">The collection of formulas to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the formulas collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
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
        /// Adds an autofilter to the specified range in the current sheet.
        /// </summary>
        /// <remarks>
        /// <para>The range height must be 1.</para>
        /// <para>Only one filter per sheet is allowed.</para>
        /// </remarks>
        /// <param name="range">The range where the autofilter should be applied.</param>
        /// <param name="overwrite">If set to <c>true</c>, any existing autofilter will be replaced.</param>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the autofilter to.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="SheetAlreadyHasFilterException">Thrown when there is already an autofilter in the current sheet and <paramref name="overwrite"/> is <c>false</c>.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when the height of the <paramref name="range"/> is not 1.</exception>
        public void AddAutofilter(string range, bool overwrite = false)
            => AddAutofilter(new CellRange(range), overwrite);

        /// <summary>
        /// Adds an autofilter to the specified range in the current sheet.
        /// </summary>
        /// <remarks>
        /// <para>The range height must be 1.</para>
        /// <para>Only one filter per sheet is allowed.</para>
        /// </remarks>
        /// <param name="range">The range where the autofilter should be applied.</param>
        /// <param name="overwrite">If set to <c>true</c>, any existing autofilter will be replaced.</param>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the autofilter to.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="SheetAlreadyHasFilterException">Thrown when there is already an autofilter in the current sheet and <paramref name="overwrite"/> is <c>false</c>.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when the height of the <paramref name="range"/> is not 1.</exception>
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
        /// Adds a list data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="formula">The formula defining the list of valid values.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are considered valid.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the autofilter to.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
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
        /// Adds a list data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="formula">The formula defining the list of valid values.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are considered valid.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the autofilter to.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
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
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
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
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
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
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        [CLSCompliant(false)]
        public void AddIntegerValidator(string range,
                                        uint firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        uint? secondOperand = null)
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
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        [CLSCompliant(false)]
        public void AddIntegerValidator(CellRange range,
                                        uint firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        uint? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Whole, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public void AddIntegerValidator(string range,
                                        long firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        long? secondOperand = null)
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
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        public void AddIntegerValidator(CellRange range,
                                        long firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        long? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Whole, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        [CLSCompliant(false)]
        public void AddIntegerValidator(string range,
                                        ulong firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        ulong? secondOperand = null)
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
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        [CLSCompliant(false)]
        public void AddIntegerValidator(CellRange range,
                                        ulong firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        ulong? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Whole, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
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
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
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
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public void AddDecimalValidator(string range,
                                        float firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        float? secondOperand = null)
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
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        public void AddDecimalValidator(CellRange range,
                                        float firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        float? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Decimal, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public void AddDecimalValidator(string range,
                                        double firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        double? secondOperand = null)
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
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        public void AddDecimalValidator(CellRange range,
                                        double firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        double? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Decimal, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds a comment to a specified cell range.
        /// </summary>
        /// <param name="text">The text of the comment.</param>
        /// <param name="reference">The cell range where the comment will be added. Must be a single cell range.</param>
        /// <param name="author">The author of the comment. Default is "BigExcelCreator".</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="author"/> is null or empty, or when <paramref name="reference"/> is not a single cell range.</exception>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="reference"/> is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the comment to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="reference"/> does not represent a valid range.</exception>
        public void Comment(string text, string reference, string author = "BigExcelCreator")
        {
            CellRange cellRange = new(reference);
            Comment(text, cellRange, author);
        }

        /// <summary>
        /// Adds a comment to a specified cell range.
        /// </summary>
        /// <param name="text">The text of the comment.</param>
        /// <param name="cellRange">The cell range where the comment will be added. Must be a single cell range.</param>
        /// <param name="author">The author of the comment. Default is "BigExcelCreator".</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="author"/> is null or empty, or when <paramref name="cellRange"/> is not a single cell range.</exception>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="cellRange"/> is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the comment to.</exception>
        public void Comment(string text, CellRange cellRange, string author = "BigExcelCreator")
        {
            if (string.IsNullOrEmpty(author)) { throw new ArgumentOutOfRangeException(nameof(author)); }
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(cellRange);
#else
            if (cellRange == null) { throw new ArgumentNullException(nameof(cellRange)); }
#endif
            if (!cellRange.IsSingleCellRange) { throw new ArgumentOutOfRangeException(nameof(cellRange), string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoWordsConcatenation, nameof(cellRange), ConstantsAndTexts.MustBeASingleCellRange)); }
            if (!sheetOpen) { throw new NoOpenSheetException(string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoWordsConcatenation, "Comments", ConstantsAndTexts.NeedToBePlacedOnSSheet)); }

            commentManager ??= new();
            commentManager.Add(new CommentReference()
            {
                Cell = cellRange.RangeStringNoSheetName,
                Text = text,
                Author = author,
            });
        }

        /// <summary>
        /// Adds a conditional formatting rule based on a formula to the specified cell range.
        /// </summary>
        /// <param name="reference">The cell range to apply the conditional formatting to.</param>
        /// <param name="formula">The formula that determines the conditional formatting rule.</param>
        /// <param name="format">The format ID of the differential format in stylesheet to apply when the condition is met. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the conditional formatting to.</exception>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="reference"/> or <paramref name="formula"/> is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is negative.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="reference"/> does not represent a valid range.</exception>
        public void AddConditionalFormattingFormula(string reference, string formula, int format)
        {
            CellRange cellRange = new(reference);
            AddConditionalFormattingFormula(cellRange, formula, format);
        }

        /// <summary>
        /// Adds a conditional formatting rule based on a formula to the specified cell range.
        /// </summary>
        /// <param name="cellRange">The cell range to apply the conditional formatting to.</param>
        /// <param name="formula">The formula that determines the conditional formatting rule.</param>
        /// <param name="format">The format ID of the differential format in stylesheet to apply when the condition is met. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the conditional formatting to.</exception>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="cellRange"/> or <paramref name="formula"/> is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is negative.</exception>
        public void AddConditionalFormattingFormula(CellRange cellRange, string formula, int format)
        {
            if (!sheetOpen) { throw new NoOpenSheetException(ConstantsAndTexts.ConditionalFormattingMustBeOnSheet); }

#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(cellRange);
#else
            if (cellRange == null) { throw new ArgumentNullException(nameof(cellRange)); }
#endif

            if (formula.IsNullOrWhiteSpace()) { throw new ArgumentNullException(nameof(formula)); }
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif

            ConditionalFormatting conditionalFormatting = new()
            {
                SequenceOfReferences = new([cellRange.RangeStringNoSheetName]),
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
        /// Adds a conditional formatting rule based on a cell value to the specified cell range.
        /// </summary>
        /// <param name="reference">The cell range to apply the conditional formatting to.</param>
        /// <param name="operator">The operator to use for the conditional formatting rule.</param>
        /// <param name="value">The value to compare the cell value against.</param>
        /// <param name="format">The format ID of the differential format in stylesheet to apply when the condition is met. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <param name="value2">The second value to compare the cell value against, used for "Between" and "NotBetween" operators.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="reference"/>, <paramref name="value"/>, or <paramref name="value2"/> (if required) is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is negative.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the conditional formatting to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="reference"/> does not represent a valid range.</exception>
        public void AddConditionalFormattingCellIs(string reference, ConditionalFormattingOperatorValues @operator, string value, int format, string value2 = null)
        {
            CellRange cellRange = new(reference);
            AddConditionalFormattingCellIs(cellRange, @operator, value, format, value2);
        }

        /// <summary>
        /// Adds a conditional formatting rule based on a cell value to the specified cell range.
        /// </summary>
        /// <param name="cellRange">The cell range to apply the conditional formatting to.</param>
        /// <param name="operator">The operator to use for the conditional formatting rule.</param>
        /// <param name="value">The value to compare the cell value against.</param>
        /// <param name="format">The format ID of the differential format in stylesheet to apply when the condition is met. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <param name="value2">The second value to compare the cell value against, used for "Between" and "NotBetween" operators.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="cellRange"/>, <paramref name="value"/>, or <paramref name="value2"/> (if required) is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is negative.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the conditional formatting to.</exception>
        public void AddConditionalFormattingCellIs(CellRange cellRange, ConditionalFormattingOperatorValues @operator, string value, int format, string value2 = null)
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(cellRange);
#else
            if (cellRange == null) { throw new ArgumentNullException(nameof(cellRange)); }
#endif

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
                SequenceOfReferences = new([cellRange.RangeStringNoSheetName]),
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
        /// Adds a conditional formatting rule to highlight duplicated values in the specified cell range.
        /// </summary>
        /// <param name="reference">The cell range to apply the conditional formatting to.</param>
        /// <param name="format">The format ID of the differential format in stylesheet to apply when the condition is met. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="reference"/> is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is negative.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the conditional formatting to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="reference"/> does not represent a valid range.</exception>
        public void AddConditionalFormattingDuplicatedValues(string reference, int format)
        {
            CellRange cellRange = new(reference);
            AddConditionalFormattingDuplicatedValues(cellRange, format);
        }

        /// <summary>
        /// Adds a conditional formatting rule to highlight duplicated values in the specified cell range.
        /// </summary>
        /// <param name="cellRange">The cell range to apply the conditional formatting to.</param>
        /// <param name="format">The format ID of the differential format in stylesheet to apply when the condition is met. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/></param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="cellRange"/> is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is negative.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the conditional formatting to.</exception>
        public void AddConditionalFormattingDuplicatedValues(CellRange cellRange, int format)
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(cellRange);
#else
            if (cellRange == null) { throw new ArgumentNullException(nameof(cellRange)); }
#endif

#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!sheetOpen) { throw new NoOpenSheetException(ConstantsAndTexts.ConditionalFormattingMustBeOnSheet); }

            ConditionalFormatting conditionalFormatting = new()
            {
                SequenceOfReferences = new([cellRange.RangeStringNoSheetName]),
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
        /// Merges the specified cell range in the current sheet.
        /// </summary>
        /// <param name="range">The cell range to merge.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="range"/> is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to merge the cells into.</exception>
        /// <exception cref="OverlappingRangesException">Thrown when the specified range overlaps with an existing merged range.</exception>
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
        /// Merges the specified cell range in the current sheet.
        /// </summary>
        /// <param name="range">The cell range to merge.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="range"/> is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to merge the cells into.</exception>
        /// <exception cref="OverlappingRangesException">Thrown when the specified range overlaps with an existing merged range.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public void MergeCells(string range) => MergeCells(new CellRange(range));

        /// <summary>
        /// Creates a new sheet from a collection of objects, automatically mapping object properties to columns.
        /// </summary>
        /// <typeparam name="T">The type of objects in the collection. Must be a reference type.</typeparam>
        /// <param name="data">The collection of objects to write to the sheet.</param>
        /// <param name="sheetName">The name of the sheet to create.</param>
        /// <param name="writeHeaderRow">If set to <c>true</c>, writes column headers as first row. Default is <c>true</c>.</param>
        /// <param name="addAutoFilterOnFirstColumn">If set to <c>true</c>, adds an autofilter to the first row. Default is <c>false</c>.</param>
        /// <param name="columns">The column definitions to use for the sheet. If not provided, columns will be generated automatically from the object type. Default is <c>null</c>.</param>
        /// <remarks>
        /// <para>This method automatically discovers properties from type <typeparamref name="T"/> and writes them as sheet columns.</para>
        /// <para>Properties can be decorated with the following attributes to customize their behavior:</para>
        /// <list type="bullet">
        /// <item><description><see cref="ExcelIgnoreAttribute"/> - Excludes a property from being written to the sheet.</description></item>
        /// <item><description><see cref="ExcelColumnNameAttribute"/> - Sets a custom column header name.</description></item>
        /// <item><description><see cref="ExcelColumnOrderAttribute"/> - Controls the column order.</description></item>
        /// <item><description><see cref="ExcelColumnTypeAttribute"/> - Specifies the cell data type (Text, Number, or Formula).</description></item>
        /// <item><description><see cref="ExcelColumnWidthAttribute"/> - Sets a custom column width.</description></item>
        /// <item><description><see cref="ExcelColumnHiddenAttribute"/> - Hides the column from view.</description></item>
        /// <item><description><see cref="ExcelStyleNameAttribute"/> - Controls styling</description></item>
        /// <item><description><see cref="ExcelHeaderStyleNameAttribute"/> (in class) - Controls header row styling</description></item>
        /// </list>
        /// <para>The sheet state is set to <see cref="SheetStateValues.Visible"/> by default.</para>
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="data"/> is <c>null</c>.</exception>
        /// <exception cref="SheetAlreadyOpenException">Thrown when a sheet is already open and not closed before opening a new one.</exception>
        /// <exception cref="SheetNameCannotBeEmptyException">Thrown when <paramref name="sheetName"/> is null or empty.</exception>
        /// <exception cref="SheetWithSameNameAlreadyExistsException">Thrown when a sheet with the same name already exists.</exception>
        public void CreateSheetFromObject<T>(IEnumerable<T> data, string sheetName, bool writeHeaderRow = true, bool addAutoFilterOnFirstColumn = false, IList<Column> columns = default)
             where T : class => CreateSheetFromObject(data, sheetName, SheetStateValues.Visible, writeHeaderRow, addAutoFilterOnFirstColumn, columns);

        /// <summary>
        /// Creates a new sheet from a collection of objects with a specified sheet state, automatically mapping object properties to columns.
        /// </summary>
        /// <typeparam name="T">The type of objects in the collection. Must be a reference type.</typeparam>
        /// <param name="data">The collection of objects to write to the sheet.</param>
        /// <param name="sheetName">The name of the sheet to create.</param>
        /// <param name="sheetState">Sets sheet visibility. <c>SheetStateValues.Visible</c> to list the sheet. <c>SheetStateValues.Hidden</c> to hide it. <c>SheetStateValues.VeryHidden</c> to hide it and prevent unhiding from the GUI.</param>
        /// <param name="writeHeaderRow">If set to <c>true</c>, writes column headers as first row. Default is <c>true</c>.</param>
        /// <param name="addAutoFilterOnFirstColumn">If set to <c>true</c>, adds an autofilter to the first row. Default is <c>false</c>.</param>
        /// <param name="columns">The column definitions to use for the sheet. If not provided, columns will be generated automatically from the object type. Default is <c>null</c>.</param>
        /// <remarks>
        /// <para>This method automatically discovers properties from type <typeparamref name="T"/> and writes them as sheet columns.</para>
        /// <para>Properties can be decorated with the following attributes to customize their behavior:</para>
        /// <list type="bullet">
        /// <item><description><see cref="ExcelIgnoreAttribute"/> - Excludes a property from being written to the sheet.</description></item>
        /// <item><description><see cref="ExcelColumnNameAttribute"/> - Sets a custom column header name.</description></item>
        /// <item><description><see cref="ExcelColumnOrderAttribute"/> - Controls the column order.</description></item>
        /// <item><description><see cref="ExcelColumnTypeAttribute"/> - Specifies the cell data type (Text, Number, or Formula).</description></item>
        /// <item><description><see cref="ExcelColumnWidthAttribute"/> - Sets a custom column width.</description></item>
        /// <item><description><see cref="ExcelColumnHiddenAttribute"/> - Hides the column from view.</description></item>
        /// <item><description><see cref="ExcelStyleNameAttribute"/> - Controls styling</description></item>
        /// <item><description><see cref="ExcelHeaderStyleNameAttribute"/> (in class) - Controls header row styling</description></item>
        /// </list>
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="data"/> is <c>null</c>.</exception>
        /// <exception cref="SheetAlreadyOpenException">Thrown when a sheet is already open and not closed before opening a new one.</exception>
        /// <exception cref="SheetNameCannotBeEmptyException">Thrown when <paramref name="sheetName"/> is null or empty.</exception>
        /// <exception cref="SheetWithSameNameAlreadyExistsException">Thrown when a sheet with the same name already exists.</exception>
        public void CreateSheetFromObject<T>(IEnumerable<T> data, string sheetName, SheetStateValues sheetState, bool writeHeaderRow = true, bool addAutoFilterOnFirstColumn = false, IList<Column> columns = default)
            where T : class
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(data);
#else
            if (data is null) throw new ArgumentNullException(nameof(data));
#endif

            if (columns?.Any() != true) { columns = CreateColumnsFromObject(typeof(T)); }

            CreateAndOpenSheet(sheetName, columns, sheetState);

            IOrderedEnumerable<PropertyInfo> sortedColumns = GetColumnsOrdered(typeof(T));

            ExcelHeaderStyleNameAttribute headerStyle = typeof(T)
                .GetCustomAttributes(typeof(ExcelHeaderStyleNameAttribute), false)
                .Cast<ExcelHeaderStyleNameAttribute>()
                .FirstOrDefault();

            if (writeHeaderRow)
            {
                writeHeaderRowFromData(sortedColumns, headerStyle);
            }

            if (addAutoFilterOnFirstColumn)
            {
                CellRange autoFilterRange = new(1, 1, sortedColumns.Count(), 1, sheetName);
                AddAutofilter(autoFilterRange);
            }

            foreach (T dataRow in data)
            {
                BeginRow();
                foreach (PropertyInfo columnName in sortedColumns)
                {
                    int cellFormat =
                        columnName.GetCustomAttributes(typeof(ExcelStyleNameAttribute), false)
                        .Cast<ExcelStyleNameAttribute>()
                        .FirstOrDefault()?
                        .Format ?? 0;
                    CellDataType cellType =
                        columnName.GetCustomAttributes(typeof(ExcelColumnTypeAttribute), false)
                        .Cast<ExcelColumnTypeAttribute>()
                        .FirstOrDefault()?
                        .Type ?? CellDataType.Text;
                    object cellData = columnName.GetValue(dataRow, null);

                    WriteCellFromData(cellData, cellType, cellFormat);
                }
                EndRow();
            }
            CloseSheet();
        }

        /// <summary>
        /// Closes the current document, ensuring all data is written and resources are released.
        /// </summary>
        /// <remarks>
        /// This method will end any open rows and sheets, write shared strings and sheets, and save the document and worksheet part writer.
        /// If saving to a stream, it will reset the stream position to the beginning.
        /// </remarks>
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
        private bool disposed = true; // There is a possibility that an exception is thrown in the constructor, so we set this to true to avoid a NullReferenceException in the finalizer.

        /// <summary>
        /// Closes the current document, ensuring all data is written and resources are released.
        /// </summary>
        /// <remarks>
        /// This method will end any open rows and sheets, write shared strings and sheets, and save the document and worksheet part writer.
        /// If saving to a stream, it will reset the stream position to the beginning.
        /// </remarks>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                // Ensure the document is properly closed.
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
        /// Closes the current document, ensuring all data is written and resources are released.
        /// </summary>
        /// <remarks>
        /// This method will end any open rows and sheets, write shared strings and sheets, and save the document and worksheet part writer.
        /// If saving to a stream, it will reset the stream position to the beginning.
        /// </remarks>
        public void Dispose() // Implement IDisposable
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Finalizes an instance of the <see cref="BigExcelWriter"/> class.
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
            if (!sheetOpen) { throw new NoOpenSheetException(string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoWordsConcatenation, "Validators", ConstantsAndTexts.NeedToBePlacedOnSSheet)); }

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

        private void WriteNumberCellInternal(string number, int format = 0)
        {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!rowOpen) { throw new NoOpenRowException(ConstantsAndTexts.NoActiveRow); }

            //reset the list of attributes
            List<OpenXmlAttribute> attributes =
            [
                //add the cell reference attribute
                new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
                //styles
                new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
            ];

            //write the cell start element with the type and reference attributes
            workSheetPartWriter.WriteStartElement(new Cell(), attributes);
            //write the cell value
            workSheetPartWriter.WriteElement(new CellValue(number));

            // write the end cell element
            workSheetPartWriter.WriteEndElement();

            columnNum++;
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
                    foreach (OpenXmlElement item in conditionalFormattingRule.ChildElements)
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

        private void SetSheetDefaults()
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

                worksheet.SheetViews = new(sheetView);
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

        private static void ThrowIfInvalidSpreadsheetDocumentType(SpreadsheetDocumentType spreadsheetDocumentType)
        {
            SpreadsheetDocumentType[] validSpreadsheetDocumentTypes =
            [
                SpreadsheetDocumentType.Workbook,
                SpreadsheetDocumentType.Template,
                SpreadsheetDocumentType.MacroEnabledWorkbook,
                SpreadsheetDocumentType.MacroEnabledTemplate,
            ];
            if (!validSpreadsheetDocumentTypes.Contains(spreadsheetDocumentType))
            {
                throw new UnsupportedSpreadsheetDocumentTypeException(string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.InvalidSpreadsheetDocumentType, spreadsheetDocumentType));
            }
        }

        private static List<Column> CreateColumnsFromObject(Type type)
        {
            List<Column> columns = [];

            var dtoCols = GetColumnsOrdered(type);
            foreach (var dtoCol in dtoCols)
            {
                bool hidden = Attribute.IsDefined(dtoCol, typeof(ExcelColumnHiddenAttribute));
                int? width = dtoCol.GetCustomAttributes(typeof(ExcelColumnWidthAttribute), false)
                  .Cast<ExcelColumnWidthAttribute>()
                  .FirstOrDefault()?
                  .Width;
                bool customWidth = width != null;
                Column column = new() { CustomWidth = customWidth, Hidden = hidden };
                if (customWidth)
                {
                    column.Width = width;
                }

                columns.Add(column);
            }

            return columns;
        }

        private static IOrderedEnumerable<PropertyInfo> GetColumnsOrdered(Type type)
        {
            return type.GetProperties()
                    .Where(x => !Attribute.IsDefined(x, typeof(ExcelIgnoreAttribute)))
                    .OrderBy(x => x.GetCustomAttributes(typeof(ExcelColumnOrderAttribute), false)
                            .Cast<ExcelColumnOrderAttribute>()
                        .FirstOrDefault()?
                        .Order ?? int.MaxValue);
        }

        private void WriteCellFromData(object cellData, CellDataType cellType, int format)
        {
            switch (cellType)
            {
                case CellDataType.Number:
                    if (cellData != null)
                    {
                        WriteNumberCell(cellData, format);
                    }
                    else
                    {
                        WriteTextCell(string.Empty, format);
                    }
                    break;
                case CellDataType.Formula:
                    string cellDataFormula = cellData?.ToString();
                    if (!cellDataFormula.IsNullOrWhiteSpace())
                    {
                        WriteFormulaCell(cellDataFormula, format);
                    }
                    else
                    {
                        WriteTextCell(string.Empty, format);
                    }
                    break;
                case CellDataType.Text:
                    string cellDataString = cellData?.ToString() ?? "";
                    WriteTextCell(cellDataString, format);
                    break;
            }
        }

        private void writeHeaderRowFromData(IOrderedEnumerable<PropertyInfo> sortedColumns, ExcelHeaderStyleNameAttribute headerFormat)
        {
            BeginRow();
            foreach (var column in sortedColumns)
            {
                string columnName = column.GetCustomAttributes(typeof(ExcelColumnNameAttribute), false).Cast<ExcelColumnNameAttribute>().FirstOrDefault()?.Name ?? column.Name;
                int format = 0;

                ExcelStyleNameAttribute columnFormat = column.GetCustomAttributes(typeof(ExcelStyleNameAttribute), false).Cast<ExcelStyleNameAttribute>().FirstOrDefault();

                if (headerFormat != null)
                {
                    format = headerFormat.Format;
                }

                if (columnFormat != null
                    && (columnFormat.HeaderStylingPriority == StylingPriority.Data
                        || headerFormat == null
                    ))
                {
                    format = columnFormat.Format;
                }

                WriteTextCell(columnName, format);
            }
            EndRow();
        }
    }
    #endregion
}
