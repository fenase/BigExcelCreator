// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes.Interfaces;
using BigExcelCreator.Enums;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Applies an integer data validation rule to a property column.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExcelIntegerValidatorAttribute : Attribute, IExcelValidatorAttribute
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
        /// Gets the second operand for operators that require it.
        /// </summary>

        /// <summary>
        /// Initializes an integer data validation attribute.
        /// </summary>
        /// <param name="firstOperand">The first operand of the validation rule.</param>
        /// <param name="validationType">The validation comparison operator.</param>
        /// <param name="allowBlank">Whether blank values are allowed.</param>
        /// <param name="showInputMessage">Whether the Excel input message is shown.</param>
        /// <param name="showErrorMessage">Whether the Excel error message is shown.</param>
        /// <param name="secondOperand">The second operand for operators that require it.</param>
        public ExcelIntegerValidatorAttribute(int firstOperand,
                                              ValidatorOperator validationType,
                                              bool allowBlank = true,
                                              bool showInputMessage = true,
                                              bool showErrorMessage = true,
                                              int secondOperand = 0)
        {
            FirstOperand = firstOperand;
            SecondOperand = secondOperand;
            ValidationType = validationType;
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
        }

        /// <summary>Initializes an unsigned 32-bit integer data validation attribute.</summary>
        [CLSCompliant(false)]
        public ExcelIntegerValidatorAttribute(uint firstOperand,
                                              ValidatorOperator validationType,
                                              bool allowBlank = true,
                                              bool showInputMessage = true,
                                              bool showErrorMessage = true,
                                              uint secondOperand = 0)
        {
            FirstOperand = firstOperand;
            SecondOperand = secondOperand;
            ValidationType = validationType;
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
        }

        /// <summary>Initializes a signed 64-bit integer data validation attribute.</summary>
        public ExcelIntegerValidatorAttribute(long firstOperand,
                                              ValidatorOperator validationType,
                                              bool allowBlank = true,
                                              bool showInputMessage = true,
                                              bool showErrorMessage = true,
                                              long secondOperand = 0)
        {
            FirstOperand = firstOperand;
            SecondOperand = secondOperand;
            ValidationType = validationType;
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
        }

        /// <summary>Initializes an unsigned 64-bit integer data validation attribute.</summary>
        [CLSCompliant(false)]
        public ExcelIntegerValidatorAttribute(ulong firstOperand,
                                              ValidatorOperator validationType,
                                              bool allowBlank = true,
                                              bool showInputMessage = true,
                                              bool showErrorMessage = true,
                                              ulong secondOperand = 0)
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
