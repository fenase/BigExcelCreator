// Copyright (c) 2022-2025, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Enums;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Specifies the Excel cell type for a property when exporting to Excel.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelColumnTypeAttribute(CellDataType type) : Attribute
    {
        /// <summary>
        /// Gets the Excel cell type for the attributed property.
        /// </summary>
        public CellDataType Type { get; } = type;
    }
}
