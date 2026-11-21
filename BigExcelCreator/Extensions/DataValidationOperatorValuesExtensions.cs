// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Enums;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace BigExcelCreator.Extensions
{
    internal static class DataValidationOperatorValuesExtensions
    {
        internal static bool RequiresSecondOperand(this DataValidationOperatorValues dataValidationOperator)
        {
            if (dataValidationOperator == DataValidationOperatorValues.Between) { return true; }
            if (dataValidationOperator == DataValidationOperatorValues.NotBetween) { return true; }
            return false;
        }

        internal static DataValidationOperatorValues ToDataValidationOperatorValues(this ValidatorOperator validationOperator)
        {
            return validationOperator switch
            {
                ValidatorOperator.Equal => DataValidationOperatorValues.Equal,
                ValidatorOperator.NotEqual => DataValidationOperatorValues.NotEqual,
                ValidatorOperator.GreaterThan => DataValidationOperatorValues.GreaterThan,
                ValidatorOperator.GreaterThanOrEqual => DataValidationOperatorValues.GreaterThanOrEqual,
                ValidatorOperator.LessThan => DataValidationOperatorValues.LessThan,
                ValidatorOperator.LessThanOrEqual => DataValidationOperatorValues.LessThanOrEqual,
                ValidatorOperator.Between => DataValidationOperatorValues.Between,
                ValidatorOperator.NotBetween => DataValidationOperatorValues.NotBetween,
                _ => throw new ArgumentOutOfRangeException(nameof(validationOperator)),
            };
        }
    }
}
