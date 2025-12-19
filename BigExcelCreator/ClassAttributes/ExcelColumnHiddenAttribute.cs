// Copyright (c) 2022-2025, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Marks a property as hidden in Excel exports.
    /// When applied to a property, the corresponding column will not be visible in the generated Excel spreadsheet.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelColumnHiddenAttribute : Attribute
    {
    }
}
