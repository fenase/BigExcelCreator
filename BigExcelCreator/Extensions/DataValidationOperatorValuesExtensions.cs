// Copyright (c) 2022-2023, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Spreadsheet;

namespace BigExcelCreator.Extensions
{
    internal static class DataValidationOperatorValuesExtensions
    {
        internal static bool RequiresSecondOperand(this DataValidationOperatorValues dataValidationOperator)
        {
            return dataValidationOperator switch
            {
                DataValidationOperatorValues.Between => true,
                DataValidationOperatorValues.NotBetween => true,
                _ => false,
            };
        }
    }
}
