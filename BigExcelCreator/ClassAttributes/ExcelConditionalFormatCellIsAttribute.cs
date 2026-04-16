// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes.Interfaces;
using BigExcelCreator.Enums;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Adds a conditional formatting rule based on a cell value (CellIs)
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    [CLSCompliant(false)]
    public sealed class ExcelConditionalFormatCellIsAttribute : Attribute, IConditionalFormatAttributes
    {
        /// <summary>
        /// The format ID of the differential format in styleSheet to apply when the condition is met. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/>.
        /// </summary>
        public int Format { get; }

        /// <summary>
        /// The differential style name to apply to the cell.
        /// </summary>
        public string StyleName { get; }

        /// <summary>
        /// The operator to use for the conditional formatting rule.
        /// </summary>
        public ConditionalFormattingOperatorValues Operator { get; }

        /// <summary>
        /// The value to compare the cell value against.
        /// </summary>
        public string Value { get; }

        /// <summary>
        /// The second value to compare the cell value against, used for "Between" and "NotBetween" operators.
        /// </summary>
        public string Value2 { get; }

        internal StyleModes StyleMode { get; }
        StyleModes IConditionalFormatAttributes.StyleMode => StyleMode;

        /// <summary>
        /// Adds a conditional formatting rule based on a cell value (CellIs) with the specified format, operator, and comparison values.
        /// </summary>
        /// <param name="format">The index of the conditional format to apply when the condition is met.</param>
        /// <param name="operator">The comparison operator used to evaluate the cell value.</param>
        /// <param name="value">The first value to compare against the cell value. Cannot be null.</param>
        /// <param name="value2">The second value to compare against the cell value, used for operators that require two values. Optional;
        /// may be null for single-value operators.</param>
        public ExcelConditionalFormatCellIsAttribute(int format, ConditionalFormattingOperatorValues @operator, string value, string value2 = null)
        {
            Format = format;
            Operator = @operator;
            Value = value;
            Value2 = value2;
            StyleMode = StyleModes.Index;
        }

        /// <summary>
        /// Adds a conditional formatting rule based on a cell value (CellIs) with the specified format, operator, and comparison values.
        /// </summary>
        /// <param name="styleName">The name of the style to apply when the conditional formatting rule is met. Cannot be null or empty.</param>
        /// <param name="operator">The comparison operator used to evaluate the cell's value against the specified criteria.</param>
        /// <param name="value">The first value to compare the cell's value against, according to the specified operator. Cannot be null.</param>
        /// <param name="value2">The second value to use for comparison when the operator requires two values; otherwise, null.</param>
        public ExcelConditionalFormatCellIsAttribute(string styleName, ConditionalFormattingOperatorValues @operator, string value, string value2 = null)
        {
            StyleName = styleName;
            Operator = @operator;
            Value = value;
            Value2 = value2;
            StyleMode = StyleModes.Name;
        }
    }
}
