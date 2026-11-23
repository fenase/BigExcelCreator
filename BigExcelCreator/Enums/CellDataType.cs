// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes;

namespace BigExcelCreator.Enums
{
    /// <summary>
    /// Specifies the Excel cell type for a spreadsheet cell.
    /// </summary>
    public enum CellDataType
    {
        /// <summary>
        /// Text cell type for string values.
        /// </summary>
        /// <remarks>
        /// When writing from a model property, the value will be converted to a string using its <c>ToString()</c> method.
        /// </remarks>
        Text,

        /// <summary>
        /// Number cell type for numeric values.
        /// </summary>
        Number,

        /// <summary>
        /// Formula cell type for Excel formulas.
        /// </summary>
        Formula,

        /// <summary>
        /// The cell type for enum values
        /// </summary>
        /// <remarks>
        /// <para>When writing from a model property, the value will be converted to a string using its Display attribute.</para>
        /// <para>When writing from a model property, the value will be converted to a string using its <c>ToString()</c> method if no Display attribute is defined.</para>
        /// <para>No validation will be applied to the value by default. Use the <see cref="ExcelListValidatorAttribute"/> to apply validation.</para>
        /// </remarks>
        EnumText,
    }
}
