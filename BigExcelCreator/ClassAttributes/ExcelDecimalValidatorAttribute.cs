// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes.Interfaces;
using BigExcelCreator.Enums;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Applies a decimal data validation rule to a property column.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelDecimalValidatorAttribute : Attribute, IExcelValidatorAttribute
    {
        /// <summary>
        /// Gets the first operand of the validation rule.
        /// </summary>

        /// <summary>
        /// Gets the validation comparison operator.
        /// </summary>
        public ValidatorOperator ValidationType { get; }

        /// <summary>Gets the first operand supplied to the attribute.</summary>
        public object FirstOperand { get; }

        /// <summary>Gets the second operand supplied to the attribute.</summary>
        public object SecondOperand { get; }

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
        /// Initializes a decimal data validation attribute using a single-precision operand.
        /// </summary>
        public ExcelDecimalValidatorAttribute(float firstOperand,
                                              ValidatorOperator validationType,
                                              bool allowBlank = true,
                                              bool showInputMessage = true,
                                              bool showErrorMessage = true,
                                              float secondOperand = 0)
        {
            FirstOperand = firstOperand;
            SecondOperand = secondOperand;
            ValidationType = validationType;
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
        }

        /// <summary>
        /// Initializes a decimal data validation attribute using a double-precision operand.
        /// </summary>
        public ExcelDecimalValidatorAttribute(double firstOperand,
                                              ValidatorOperator validationType,
                                              bool allowBlank = true,
                                              bool showInputMessage = true,
                                              bool showErrorMessage = true,
                                              double secondOperand = 0)
        {
            FirstOperand = firstOperand;
            SecondOperand = secondOperand;
            ValidationType = validationType;
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
        }
    }
}
