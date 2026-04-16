// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.Enums;
using BigExcelCreator.Exceptions;
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
    /// </summary>
    /// <remarks>
    /// <para><see href="https://www.nuget.org/packages/BigExcelCreator">NuGet</see></para>
    /// <para><seealso href="https://github.com/fenase/BigExcelCreator">Source</seealso></para>
    /// <para><seealso href="https://fenase.github.io/BigExcelCreator/api/BigExcelCreator.BigExcelWriter.html">API</seealso></para>
    /// <para><seealso href="https://fenase.github.io/projects/BigExcelCreator">Site</seealso></para>
    /// </remarks>
    public partial class BigExcelWriter : IDisposable
    {
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
    }
}
