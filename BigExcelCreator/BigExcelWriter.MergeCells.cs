// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.Exceptions;
using BigExcelCreator.Extensions;
using BigExcelCreator.Ranges;
using System;

namespace BigExcelCreator
{
    public partial class BigExcelWriter : IDisposable
    {
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
    }
}
