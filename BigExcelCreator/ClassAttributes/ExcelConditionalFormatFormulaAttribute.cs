// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes.Interfaces;
using BigExcelCreator.Enums;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Adds a conditional formatting rule based on a formula
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelConditionalFormatFormulaAttribute : Attribute, IConditionalFormatAttributes
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
        /// The formula that determines the conditional formatting rule.
        /// </summary>
        public string Formula { get; }

        internal StyleModes StyleMode { get; }
        StyleModes IConditionalFormatAttributes.StyleMode => StyleMode;

        /// <summary>
        /// Adds a conditional formatting rule based on a formula
        /// </summary>
        /// <param name="format">The format ID of the differential format in styleSheet to apply when the condition is met. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/>.</param>
        /// <param name="formula">The formula that determines the conditional formatting rule.</param>
        public ExcelConditionalFormatFormulaAttribute(int format, string formula)
        {
            Format = format;
            Formula = formula;
            StyleMode = StyleModes.Index;
        }

        /// <summary>
        /// Adds a conditional formatting rule based on a formula
        /// </summary>
        /// <param name="styleName">The differential style name to apply to the cell.</param>
        /// <param name="formula">The formula that determines the conditional formatting rule.</param>
        public ExcelConditionalFormatFormulaAttribute(string styleName, string formula)
        {
            StyleName = styleName;
            Formula = formula;
            StyleMode = StyleModes.Name;
        }
    }
}
