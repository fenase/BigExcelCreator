// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes.Interfaces;
using BigExcelCreator.Enums;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Adds a conditional formatting rule to highlight duplicated values
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelConditionalFormatDuplicatedValuesAttribute : Attribute, IConditionalFormatAttributes
    {
        /// <summary>
        /// The format ID of the differential format in styleSheet to apply when the condition is met. See <see cref="Styles.StyleList.GetIndexDifferentialByName(string)"/>.
        /// </summary>
        public int Format { get; }

        /// <summary>
        /// The differential style name to apply to the cell.
        /// </summary>
        public string StyleName { get; }

        internal StyleModes StyleMode { get; }
        StyleModes IConditionalFormatAttributes.StyleMode => StyleMode;

        /// <summary>
        /// Adds a conditional formatting rule to highlight duplicated values with the specified format.
        /// </summary>
        /// <param name="format">The index of the conditional format to apply when the condition is met.</param>
        public ExcelConditionalFormatDuplicatedValuesAttribute(int format)
        {
            Format = format;
            StyleMode = StyleModes.Index;
        }

        /// <summary>
        /// Adds a conditional formatting rule to highlight duplicated values with the specified format.
        /// </summary>
        /// <param name="styleName">The name of the style to apply when the conditional formatting rule is met. Cannot be null or empty.</param>
        public ExcelConditionalFormatDuplicatedValuesAttribute(string styleName)
        {
            StyleName = styleName;
            StyleMode = StyleModes.Name;
        }
    }
}
