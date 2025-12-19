// Copyright (c) 2022-2025, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Specifies the column order for a property when exporting to Excel.
    /// </summary>
    /// <remarks>
    /// This attribute is used to control the order in which properties are exported
    /// as columns in an Excel worksheet. Properties are ordered by their <see cref="Order"/>
    /// value in ascending order.
    /// </remarks>
    /// <param name="order">The zero-based position of the column within the Excel sheet.</param>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelColumnOrderAttribute(int order) : Attribute
    {
        /// <summary>
        /// Gets the zero-based position of the item within its containing collection.
        /// </summary>
        public int Order { get; } = order;
    }
}
