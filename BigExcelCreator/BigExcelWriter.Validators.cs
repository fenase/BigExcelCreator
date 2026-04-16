// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.Exceptions;
using BigExcelCreator.Extensions;
using BigExcelCreator.Ranges;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;

namespace BigExcelCreator
{
    public partial class BigExcelWriter : IDisposable
    {
        /// <summary>
        /// Adds a list data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="formula">The formula defining the list of valid values.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are considered valid.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the autofilter to.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public void AddListValidator(string range,
                                     string formula,
                                     bool allowBlank = true,
                                     bool showInputMessage = true,
                                     bool showErrorMessage = true)
        {
            AddListValidator(new CellRange(range),
                             formula,
                             allowBlank,
                             showInputMessage,
                             showErrorMessage);
        }

        /// <summary>
        /// Adds a list data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="formula">The formula defining the list of valid values.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are considered valid.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the autofilter to.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        public void AddListValidator(CellRange range,
                                     string formula,
                                     bool allowBlank = true,
                                     bool showInputMessage = true,
                                     bool showErrorMessage = true)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.List, DataValidationOperatorValues.Equal, allowBlank, showInputMessage, showErrorMessage);

            AppendNewDataValidation(dataValidation, formula);
        }

        /// <summary>
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public void AddIntegerValidator(string range,
                                        int firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        int? secondOperand = null)
        {
            AddIntegerValidator(new CellRange(range),
                                firstOperand,
                                validationType,
                                allowBlank,
                                showInputMessage,
                                showErrorMessage,
                                secondOperand);
        }

        /// <summary>
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        public void AddIntegerValidator(CellRange range,
                                        int firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        int? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Whole, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        [CLSCompliant(false)]
        public void AddIntegerValidator(string range,
                                        uint firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        uint? secondOperand = null)
        {
            AddIntegerValidator(new CellRange(range),
                                firstOperand,
                                validationType,
                                allowBlank,
                                showInputMessage,
                                showErrorMessage,
                                secondOperand);
        }

        /// <summary>
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        [CLSCompliant(false)]
        public void AddIntegerValidator(CellRange range,
                                        uint firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        uint? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Whole, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public void AddIntegerValidator(string range,
                                        long firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        long? secondOperand = null)
        {
            AddIntegerValidator(new CellRange(range),
                                firstOperand,
                                validationType,
                                allowBlank,
                                showInputMessage,
                                showErrorMessage,
                                secondOperand);
        }

        /// <summary>
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        public void AddIntegerValidator(CellRange range,
                                        long firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        long? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Whole, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        [CLSCompliant(false)]
        public void AddIntegerValidator(string range,
                                        ulong firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        ulong? secondOperand = null)
        {
            AddIntegerValidator(new CellRange(range),
                                firstOperand,
                                validationType,
                                allowBlank,
                                showInputMessage,
                                showErrorMessage,
                                secondOperand);
        }

        /// <summary>
        /// Adds an integer data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        [CLSCompliant(false)]
        public void AddIntegerValidator(CellRange range,
                                        ulong firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        ulong? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Whole, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public void AddDecimalValidator(string range,
                                        decimal firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        decimal? secondOperand = null)
        {
            AddDecimalValidator(new CellRange(range),
                                firstOperand,
                                validationType,
                                allowBlank,
                                showInputMessage,
                                showErrorMessage,
                                secondOperand);
        }

        /// <summary>
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        public void AddDecimalValidator(CellRange range,
                                        decimal firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        decimal? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Decimal, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public void AddDecimalValidator(string range,
                                        float firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        float? secondOperand = null)
        {
            AddDecimalValidator(new CellRange(range),
                                firstOperand,
                                validationType,
                                allowBlank,
                                showInputMessage,
                                showErrorMessage,
                                secondOperand);
        }

        /// <summary>
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        public void AddDecimalValidator(CellRange range,
                                        float firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        float? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Decimal, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public void AddDecimalValidator(string range,
                                        double firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        double? secondOperand = null)
        {
            AddDecimalValidator(new CellRange(range),
                                firstOperand,
                                validationType,
                                allowBlank,
                                showInputMessage,
                                showErrorMessage,
                                secondOperand);
        }

        /// <summary>
        /// Adds a decimal data validation to the specified cell range.
        /// </summary>
        /// <param name="range">The cell range to apply the validation to.</param>
        /// <param name="firstOperand">The first operand for the validation.</param>
        /// <param name="validationType">The type of validation to apply.</param>
        /// <param name="allowBlank">If set to <c>true</c>, blank values are allowed.</param>
        /// <param name="showInputMessage">If set to <c>true</c>, an input message will be shown.</param>
        /// <param name="showErrorMessage">If set to <c>true</c>, an error message will be shown when invalid data is entered.</param>
        /// <param name="secondOperand">The second operand for the validation, if required by the validation type.</param>
        /// <exception cref="ArgumentNullException">Thrown when the validation type requires a second operand but <paramref name="secondOperand"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to add the validation to.</exception>
        public void AddDecimalValidator(CellRange range,
                                        double firstOperand,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true,
                                        double? secondOperand = null)
        {
            DataValidation dataValidation = AddValidatorCommon(range, DataValidationValues.Decimal, validationType, allowBlank, showInputMessage, showErrorMessage);

            if (validationType.RequiresSecondOperand() && secondOperand == null)
            {
                throw new ArgumentNullException(nameof(secondOperand), $"validation type {validationType} requires a second operand");
            }

            AppendNewDataValidation(dataValidation, firstOperand.ToString(CultureInfo.InvariantCulture), secondOperand?.ToString(CultureInfo.InvariantCulture));
        }
    }
}
