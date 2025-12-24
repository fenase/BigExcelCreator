// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.CommentsManager;
using BigExcelCreator.Enums;
using BigExcelCreator.Exceptions;
using BigExcelCreator.Ranges;
using BigExcelCreator.Styles;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;

namespace BigExcelCreator
{
    public partial class BigExcelWriter : IDisposable
    {
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

        private StyleList StyleList { get; }

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
    }
}
