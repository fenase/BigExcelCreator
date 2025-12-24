// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.Enums;
using BigExcelCreator.Styles;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;

namespace BigExcelCreator
{
    public partial class BigExcelWriter : IDisposable
    {
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
        /// <remarks>Initializes a new Workbook</remarks>
        /// <param name="stream">The stream to write the Excel document to.</param>
        /// <param name="styleList">The <see cref="StyleList"/> from where to get the stylesheet to apply to the Excel document. Enables writing styled data using the style names.</param>
        public BigExcelWriter(Stream stream, StyleList styleList)
                : this(stream, SpreadsheetDocumentType.Workbook, styleList) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified stream, spreadsheet document type, and stylesheet.
        /// </summary>
        /// <param name="stream">The stream to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="stylesheet">The stylesheet to apply to the Excel document. See <see cref="Styles.StyleList.GetStylesheet()"/>.</param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
                : this(stream, spreadsheetDocumentType, false, stylesheet) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified stream, spreadsheet document type, and stylesheet.
        /// </summary>
        /// <param name="stream">The stream to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="styleList">The <see cref="StyleList"/> from where to get the stylesheet to apply to the Excel document. Enables writing styled data using the style names.</param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, StyleList styleList)
                : this(stream, spreadsheetDocumentType, false, styleList) { }

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
        /// <remarks>Initializes a new Workbook</remarks>
        /// </summary>
        /// <param name="path">The file path to write the Excel document to.</param>
        /// <param name="styleList">The <see cref="StyleList"/> from where to get the stylesheet to apply to the Excel document. Enables writing styled data using the style names.</param>
        public BigExcelWriter(string path, StyleList styleList)
        : this(path, SpreadsheetDocumentType.Workbook, styleList) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified file path, spreadsheet document type, and stylesheet.
        /// </summary>
        /// <param name="path">The file path to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="stylesheet">The stylesheet to apply to the Excel document. See <see cref="Styles.StyleList.GetStylesheet()"/></param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
        : this(path, spreadsheetDocumentType, false, stylesheet) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified file path, spreadsheet document type, and stylesheet.
        /// </summary>
        /// <param name="path">The file path to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="styleList">The <see cref="StyleList"/> from where to get the stylesheet to apply to the Excel document. Enables writing styled data using the style names.</param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, StyleList styleList)
        : this(path, spreadsheetDocumentType, false, styleList) { }

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
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified file path, spreadsheet document type, a flag indicating whether to skip cells when they are empty, and a stylesheet.
        /// </summary>
        /// <param name="path">The file path to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="skipCellWhenEmpty">A flag indicating whether to skip cells when they are empty. When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        /// <param name="styleList">The <see cref="StyleList"/> from where to get the stylesheet to apply to the Excel document. Enables writing styled data using the style names.</param>
        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, StyleList styleList)
        {
            ThrowIfInvalidSpreadsheetDocumentType(spreadsheetDocumentType);
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(styleList);
#else
            if (styleList == null) { throw new ArgumentNullException(nameof(styleList)); }
#endif
            Path = path;
            SavingTo = SavingTo.file;
            Document = SpreadsheetDocument.Create(Path, spreadsheetDocumentType);
            StyleList = styleList;
            AllowedStyleModes |= StyleModes.Name;
            Stylesheet stylesheet = styleList.GetStylesheet();
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

        /// <summary>
        /// Initializes a new instance of the <see cref="BigExcelWriter"/> class with the specified stream, spreadsheet document type, a flag indicating whether to skip cells when they are empty, and a stylesheet.
        /// </summary>
        /// <param name="stream">The stream to write the Excel document to.</param>
        /// <param name="spreadsheetDocumentType">The type of the spreadsheet document (e.g., Workbook, Template).</param>
        /// <param name="skipCellWhenEmpty">A flag indicating whether to skip cells when they are empty. When <see langword="true"/>, writing an empty value to a cell moves the next cell to be written. When <see langword="false"/>, writing an empty value to a cell does nothing.</param>
        /// <param name="styleList">The <see cref="StyleList"/> from where to get the stylesheet to apply to the Excel document. Enables writing styled data using the style names.</param>
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, StyleList styleList)
        {
            ThrowIfInvalidSpreadsheetDocumentType(spreadsheetDocumentType);
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(styleList);
#else
            if (styleList == null) { throw new ArgumentNullException(nameof(styleList)); }
#endif
            Stream = stream;
            SavingTo = SavingTo.stream;
            Document = SpreadsheetDocument.Create(Stream, spreadsheetDocumentType);
            StyleList = styleList;
            AllowedStyleModes |= StyleModes.Name;
            Stylesheet stylesheet = styleList.GetStylesheet();
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
                AllowedStyleModes |= StyleModes.Index;
            }

            SharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();

            SkipCellWhenEmpty = skipCellWhenEmpty;
        }
    }
}
