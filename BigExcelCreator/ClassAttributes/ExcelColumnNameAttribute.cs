// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Specifies the name of an Excel column for a property.
    /// </summary>
    /// <param name="name">The name of the Excel column.</param>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelColumnNameAttribute(string name) : Attribute
    {
        /// <summary>
        /// Gets the name of the Excel column.
        /// </summary>
        public string Name { get; } = name;
    }
}
