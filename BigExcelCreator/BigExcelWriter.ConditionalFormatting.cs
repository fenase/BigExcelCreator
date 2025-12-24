// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.Exceptions;
using BigExcelCreator.Extensions;
using BigExcelCreator.Ranges;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace BigExcelCreator
{
    public partial class BigExcelWriter : IDisposable
    {
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
    }
}
