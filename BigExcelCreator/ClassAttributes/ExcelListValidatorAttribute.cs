// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes.Interfaces;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Applies a literal list data validation to a property column.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    [CLSCompliant(false)]
    public sealed class ExcelListValidatorAttribute : Attribute, IExcelValidatorAttribute
    {
        /// <summary>
        /// Gets the literal values used by the list data validation.
        /// </summary>
        public string[] Values { get; }

        /// <summary>
        /// Gets whether blank values are allowed.
        /// </summary>
        public bool AllowBlank { get; }

        /// <summary>
        /// Gets whether the Excel input message is shown.
        /// </summary>
        public bool ShowInputMessage { get; }

        /// <summary>
        /// Gets whether the Excel error message is shown.
        /// </summary>
        public bool ShowErrorMessage { get; }

        /// <summary>
        /// Initializes a literal list data validation attribute.
        /// </summary>
        /// <param name="values">The literal values accepted by the validation.</param>
        public ExcelListValidatorAttribute(params string[] values)
            : this(true, true, true, values)
        {
        }

        /// <summary>
        /// Initializes a literal list data validation attribute with display options.
        /// </summary>
        /// <param name="allowBlank">Whether blank values are allowed.</param>
        /// <param name="showInputMessage">Whether the Excel input message is shown.</param>
        /// <param name="showErrorMessage">Whether the Excel error message is shown.</param>
        /// <param name="values">The literal values accepted by the validation.</param>
        public ExcelListValidatorAttribute(bool allowBlank, bool showInputMessage, bool showErrorMessage, params string[] values)
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(values);
#else
            if (values == null)
            {
                throw new ArgumentNullException(nameof(values));
            }
#endif

            Values = values;
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
        }
    }
}
