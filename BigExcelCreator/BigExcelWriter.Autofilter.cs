// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.Exceptions;
using BigExcelCreator.Ranges;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace BigExcelCreator
{
    public partial class BigExcelWriter : IDisposable
    {
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
    }
}
