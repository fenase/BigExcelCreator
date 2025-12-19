// Copyright (c) 2022-2025, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

namespace BigExcelCreator.Enums
{
    /// <summary>
    /// Specifies the Excel cell type for a spreadsheet cell.
    /// </summary>
    public enum CellType
    {
        /// <summary>
        /// Text cell type for string values.
        /// </summary>
        Text,

        /// <summary>
        /// Number cell type for numeric values.
        /// </summary>
        Number,

        /// <summary>
        /// Formula cell type for Excel formulas.
        /// </summary>
        Formula,
    }
}
