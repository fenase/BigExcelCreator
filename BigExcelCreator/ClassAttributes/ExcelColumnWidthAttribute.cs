// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Specifies the width of an Excel column for a property.
    /// </summary>
    /// <remarks>
    /// This attribute is used to define the column width when exporting data to Excel.
    /// The width value must be a non-negative integer.
    /// </remarks>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelColumnWidthAttribute : Attribute
    {
        /// <summary>
        /// Gets the width of the Excel column.
        /// </summary>
        public int Width { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColumnWidthAttribute"/> class.
        /// </summary>
        /// <param name="width">The width of the Excel column. Must be non-negative.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="width"/> is negative.</exception>
        public ExcelColumnWidthAttribute(int width)
        {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(width);
#else
            if (width < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
#endif
            Width = width;
        }
    }
}
