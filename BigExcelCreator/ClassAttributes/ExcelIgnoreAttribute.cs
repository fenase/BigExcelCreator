// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Marks a property to be ignored when generating Excel output.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelIgnoreAttribute : Attribute { }
}
