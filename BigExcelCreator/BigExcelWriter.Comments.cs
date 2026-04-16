// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.CommentsManager;
using BigExcelCreator.Exceptions;
using BigExcelCreator.Extensions;
using BigExcelCreator.Ranges;
using System;
using System.Globalization;

namespace BigExcelCreator
{
    public partial class BigExcelWriter : IDisposable
    {
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
    }
}
