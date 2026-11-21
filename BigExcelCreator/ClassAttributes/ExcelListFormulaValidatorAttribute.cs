// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes.Interfaces;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Applies a formula-based list data validation to a property column.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelListFormulaValidatorAttribute : Attribute, IExcelValidatorAttribute
    {
        /// <summary>
        /// Gets the formula that supplies valid list values.
        /// </summary>
        public string Formula { get; }

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
        /// Initializes a formula-based list data validation attribute.
        /// </summary>
        /// <param name="formula">The formula that supplies valid list values.</param>
        public ExcelListFormulaValidatorAttribute(string formula)
            : this(formula, true, true, true)
        {
        }

        /// <summary>
        /// Initializes a formula-based list data validation attribute with display options.
        /// </summary>
        /// <param name="formula">The formula that supplies valid list values.</param>
        /// <param name="allowBlank">Whether blank values are allowed.</param>
        /// <param name="showInputMessage">Whether the Excel input message is shown.</param>
        /// <param name="showErrorMessage">Whether the Excel error message is shown.</param>
        public ExcelListFormulaValidatorAttribute(string formula, bool allowBlank, bool showInputMessage, bool showErrorMessage)
        {
            Formula = formula ?? throw new ArgumentNullException(nameof(formula));
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
        }
    }
}
