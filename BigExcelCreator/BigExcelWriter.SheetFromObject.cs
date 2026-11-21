// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.ClassAttributes;
using BigExcelCreator.ClassAttributes.Interfaces;
using BigExcelCreator.Enums;
using BigExcelCreator.Exceptions;
using BigExcelCreator.Extensions;
using BigExcelCreator.Ranges;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace BigExcelCreator
{
    public partial class BigExcelWriter : IDisposable
    {
        /// <summary>
        /// Creates a new sheet from a collection of objects, automatically mapping object properties to columns.
        /// </summary>
        /// <typeparam name="T">The type of objects in the collection. Must be a reference type.</typeparam>
        /// <param name="data">The collection of objects to write to the sheet.</param>
        /// <param name="sheetName">The name of the sheet to create.</param>
        /// <param name="writeHeaderRow">If set to <c>true</c>, writes column headers as first row. Default is <c>true</c>.</param>
        /// <param name="addAutoFilterOnFirstColumn">If set to <c>true</c>, adds an autofilter to the first row. Default is <c>false</c>.</param>
        /// <param name="columns">The column definitions to use for the sheet. If not provided, columns will be generated automatically from the object type. Default is <c>null</c>.</param>
        /// <param name="useSharedStringsOnTextData">Indicates whether to write the text values to the shared strings table. This might help reduce the output file size when the same text is shared multiple times among sheets. Default is false.</param>
        /// <param name="keepSheetOpen">Indicates whether to keep the sheet open after writing the data. Default is false. Useful for further modifications, like adding data validation. You MUST call <see cref="CloseSheet"/> manually if this is set to <c>true</c>.</param>
        /// <remarks>
        /// <para>This method automatically discovers properties from type <typeparamref name="T"/> and writes them as sheet columns.</para>
        /// <para>Properties can be decorated with the following attributes to customize their behavior:</para>
        /// <list type="bullet">
        /// <item><description><see cref="ExcelIgnoreAttribute"/> - Excludes a property from being written to the sheet.</description></item>
        /// <item><description><see cref="ExcelColumnNameAttribute"/> - Sets a custom column header name.</description></item>
        /// <item><description><see cref="ExcelColumnOrderAttribute"/> - Controls the column order.</description></item>
        /// <item><description><see cref="ExcelColumnTypeAttribute"/> - Specifies the cell data type (Text, Number, or Formula).</description></item>
        /// <item><description><see cref="ExcelColumnWidthAttribute"/> - Sets a custom column width.</description></item>
        /// <item><description><see cref="ExcelColumnHiddenAttribute"/> - Hides the column from view.</description></item>
        /// <item><description><see cref="ExcelStyleFormatAttribute"/> - Controls styling</description></item>
        /// <item><description><see cref="ExcelHeaderStyleFormatAttribute"/> (in class) - Controls header row styling</description></item>
        /// <item><description><see cref="ExcelListValidatorAttribute"/> - Applies a literal list validation.</description></item>
        /// <item><description><see cref="ExcelListFormulaValidatorAttribute"/> - Applies a formula-based list validation.</description></item>
        /// <item><description><see cref="ExcelIntegerValidatorAttribute"/> - Applies an integer validation.</description></item>
        /// <item><description><see cref="ExcelDecimalValidatorAttribute"/> - Applies a decimal validation.</description></item>
        /// </list>
        /// <para>The sheet state is set to <see cref="SheetStateValues.Visible"/> by default.</para>
        /// <para>
        /// If <paramref name="keepSheetOpen"/> is set to <c>true</c>, the sheet will remain open after writing the data, allowing for further modifications. Otherwise, the sheet will be closed automatically.
        /// You MUST call <see cref="CloseSheet"/> manually if <paramref name="keepSheetOpen"/> is set to <c>true</c>.
        /// </para>
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="data"/> is <c>null</c>.</exception>
        /// <exception cref="SheetAlreadyOpenException">Thrown when a sheet is already open and not closed before opening a new one.</exception>
        /// <exception cref="SheetNameCannotBeEmptyException">Thrown when <paramref name="sheetName"/> is null or empty.</exception>
        /// <exception cref="SheetWithSameNameAlreadyExistsException">Thrown when a sheet with the same name already exists.</exception>
        public void CreateSheetFromObject<T>(IEnumerable<T> data, string sheetName, bool writeHeaderRow = true, bool addAutoFilterOnFirstColumn = false, IList<Column> columns = default, bool useSharedStringsOnTextData = false, bool keepSheetOpen = false)
             where T : class => CreateSheetFromObject(data, sheetName, SheetStateValues.Visible, writeHeaderRow, addAutoFilterOnFirstColumn, columns, useSharedStringsOnTextData, keepSheetOpen);

        /// <summary>
        /// Creates a new sheet from a collection of objects with a specified sheet state, automatically mapping object properties to columns.
        /// </summary>
        /// <typeparam name="T">The type of objects in the collection. Must be a reference type.</typeparam>
        /// <param name="data">The collection of objects to write to the sheet.</param>
        /// <param name="sheetName">The name of the sheet to create.</param>
        /// <param name="sheetState">Sets sheet visibility. <c>SheetStateValues.Visible</c> to list the sheet. <c>SheetStateValues.Hidden</c> to hide it. <c>SheetStateValues.VeryHidden</c> to hide it and prevent unhiding from the GUI.</param>
        /// <param name="writeHeaderRow">If set to <c>true</c>, writes column headers as first row. Default is <c>true</c>.</param>
        /// <param name="addAutoFilterOnFirstColumn">If set to <c>true</c>, adds an autofilter to the first row. Default is <c>false</c>.</param>
        /// <param name="columns">The column definitions to use for the sheet. If not provided, columns will be generated automatically from the object type. Default is <c>null</c>.</param>
        /// <param name="useSharedStringsOnTextData">Indicates whether to write the text values to the shared strings table. This might help reduce the output file size when the same text is shared multiple times among sheets. Default is false.</param>
        /// <param name="keepSheetOpen">Indicates whether to keep the sheet open after writing the data. Default is false. Useful for further modifications, like adding data validation. You MUST call <see cref="CloseSheet"/> manually if this is set to <c>true</c>.</param>
        /// <remarks>
        /// <para>This method automatically discovers properties from type <typeparamref name="T"/> and writes them as sheet columns.</para>
        /// <para>Properties can be decorated with the following attributes to customize their behavior:</para>
        /// <list type="bullet">
        /// <item><description><see cref="ExcelIgnoreAttribute"/> - Excludes a property from being written to the sheet.</description></item>
        /// <item><description><see cref="ExcelColumnNameAttribute"/> - Sets a custom column header name.</description></item>
        /// <item><description><see cref="ExcelColumnOrderAttribute"/> - Controls the column order.</description></item>
        /// <item><description><see cref="ExcelColumnTypeAttribute"/> - Specifies the cell data type (Text, Number, or Formula).</description></item>
        /// <item><description><see cref="ExcelColumnWidthAttribute"/> - Sets a custom column width.</description></item>
        /// <item><description><see cref="ExcelColumnHiddenAttribute"/> - Hides the column from view.</description></item>
        /// <item><description><see cref="ExcelStyleFormatAttribute"/> - Controls styling</description></item>
        /// <item><description><see cref="ExcelHeaderStyleFormatAttribute"/> (in class) - Controls header row styling</description></item>
        /// <item><description><see cref="ExcelListValidatorAttribute"/> - Applies a literal list validation.</description></item>
        /// <item><description><see cref="ExcelListFormulaValidatorAttribute"/> - Applies a formula-based list validation.</description></item>
        /// <item><description><see cref="ExcelIntegerValidatorAttribute"/> - Applies an integer validation.</description></item>
        /// <item><description><see cref="ExcelDecimalValidatorAttribute"/> - Applies a decimal validation.</description></item>
        /// </list>
        /// <para>
        /// If <paramref name="keepSheetOpen"/> is set to <c>true</c>, the sheet will remain open after writing the data, allowing for further modifications. Otherwise, the sheet will be closed automatically.
        /// You MUST call <see cref="CloseSheet"/> manually if <paramref name="keepSheetOpen"/> is set to <c>true</c>.
        /// </para>
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="data"/> is <c>null</c>.</exception>
        /// <exception cref="SheetAlreadyOpenException">Thrown when a sheet is already open and not closed before opening a new one.</exception>
        /// <exception cref="SheetNameCannotBeEmptyException">Thrown when <paramref name="sheetName"/> is null or empty.</exception>
        /// <exception cref="SheetWithSameNameAlreadyExistsException">Thrown when a sheet with the same name already exists.</exception>
#pragma warning disable S107
        public void CreateSheetFromObject<T>(IEnumerable<T> data,
                                             string sheetName,
                                             SheetStateValues sheetState,
                                             bool writeHeaderRow = true,
                                             bool addAutoFilterOnFirstColumn = false,
                                             IList<Column> columns = default,
                                             bool useSharedStringsOnTextData = false,
                                             bool keepSheetOpen = false)
#pragma warning restore S107
            where T : class
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(data);
#else
            if (data is null) throw new ArgumentNullException(nameof(data));
#endif
            IList<T> list = data as IList<T> ?? [.. data];

            if (!(columns?.Count > 0)) { columns = CreateSpreadsheetColumnsFromObject(typeof(T)); }

            List<PropertyInfo> sortedColumns = [.. GetOrderedColumnProperties(typeof(T))];

            Dictionary<PropertyInfo, string> literalListValidatorFormulae = CreateAuxSheetForListValidator(sortedColumns, sheetName);

            CreateAndOpenSheet(sheetName, columns, sheetState);

            ExcelHeaderStyleFormatAttribute headerStyle = typeof(T)
                .GetCustomAttributes(typeof(ExcelHeaderStyleFormatAttribute), false)
                .Cast<ExcelHeaderStyleFormatAttribute>()
                .FirstOrDefault();

            if (writeHeaderRow)
            {
                WriteHeaderRowFromData(sortedColumns, headerStyle);
            }

            if (addAutoFilterOnFirstColumn)
            {
                CellRange autoFilterRange = new(1, 1, sortedColumns.Count, 1, sheetName);
                AddAutofilter(autoFilterRange);
            }

            foreach (T dataRow in list)
            {
                BeginRow();
                foreach (PropertyInfo columnName in sortedColumns)
                {
                    ExcelStyleFormatAttribute cellFormat =
                        columnName.GetCustomAttributes(typeof(ExcelStyleFormatAttribute), false)
                        .Cast<ExcelStyleFormatAttribute>()
                        .FirstOrDefault();
                    int cellStyleIndex = GetStyleFormatIndexFromAttributeAndStyleList(cellFormat);

                    CellDataType cellType =
                        columnName.GetCustomAttributes(typeof(ExcelColumnTypeAttribute), false)
                        .Cast<ExcelColumnTypeAttribute>()
                        .FirstOrDefault()?
                        .Type ?? CellDataType.Text;

                    object cellData = columnName.GetValue(dataRow, null);

                    WriteCellFromData(cellData, cellType, cellStyleIndex, useSharedStringsOnTextData);
                }
                EndRow();
            }

            AttachConditionalFormattingToSheet(writeHeaderRow, list, sortedColumns);

            AttachValidatorsToSheet(writeHeaderRow, list, sortedColumns, literalListValidatorFormulae);

            if (!keepSheetOpen)
            {
                CloseSheet();
            }
        }

        private void AttachConditionalFormattingToSheet<T>(bool writeHeaderRow, IList<T> list, List<PropertyInfo> sortedColumns)
             where T : class
        {
            int colNum = 0;
            int baseRow = writeHeaderRow ? 2 : 1;
            foreach (PropertyInfo column in sortedColumns)
            {
                colNum++;
                IConditionalFormatAttributes styleFormat =
                    column.GetCustomAttributes(typeof(IConditionalFormatAttributes), false)
                    .Cast<IConditionalFormatAttributes>()
                    .FirstOrDefault();
                if (styleFormat == null) { continue; }

                CellRange conditionalFormatRange = new(colNum, baseRow, colNum, Math.Max(baseRow + list.Count - 1, baseRow), currentSheetName);
                int format = GetStyleDifferentialFormatIndexFromAttributeAndStyleList(styleFormat);

                switch (styleFormat)
                {
                    case ExcelConditionalFormatFormulaAttribute formulaAttribute:
                        AddConditionalFormattingFormula(conditionalFormatRange, formulaAttribute.Formula, format);
                        break;
                    case ExcelConditionalFormatCellIsAttribute cellIsAttribute:
                        AddConditionalFormattingCellIs(
                            conditionalFormatRange,
                            cellIsAttribute.Operator,
                            cellIsAttribute.Value,
                            format,
                            cellIsAttribute.Value2);
                        break;
                    case ExcelConditionalFormatDuplicatedValuesAttribute _:
                        AddConditionalFormattingDuplicatedValues(conditionalFormatRange, format);
                        break;
                    default:
                        throw new NotImplementedException();
                }
            }
        }

        private Dictionary<PropertyInfo, string> CreateAuxSheetForListValidator(List<PropertyInfo> sortedColumns, string sheetName)
        {
            Dictionary<PropertyInfo, string> literalListValidatorFormulae = [];

            foreach (PropertyInfo column in sortedColumns)
            {
                ExcelListValidatorAttribute listAttribute = column
                    .GetCustomAttributes(typeof(ExcelListValidatorAttribute), false)
                    .Cast<ExcelListValidatorAttribute>()
                    .FirstOrDefault();

                if (listAttribute == null)
                {
                    continue;
                }

                string hiddenSheetName = sheetName.CreateValidatorListSheetName(column.Name);
                literalListValidatorFormulae[column] = CreateLiteralListValidationSheet(hiddenSheetName, listAttribute.Values);
            }

            return literalListValidatorFormulae;
        }

        private void AttachValidatorsToSheet<T>(bool writeHeaderRow, IList<T> list, List<PropertyInfo> sortedColumns, Dictionary<PropertyInfo, string> literalListValidatorFormulae)
            where T : class
        {
            int validationBaseRow = writeHeaderRow ? 2 : 1;
            int validationEndRow = Math.Max(validationBaseRow + list.Count - 1, validationBaseRow);

            foreach (PropertyInfo column in sortedColumns)
            {
                int columnNumber = sortedColumns.IndexOf(column) + 1;
                CellRange validationRange = new(columnNumber, validationBaseRow, columnNumber, validationEndRow, null);

                ExcelListValidatorAttribute listAttribute = column
                    .GetCustomAttributes(typeof(ExcelListValidatorAttribute), false)
                    .Cast<ExcelListValidatorAttribute>()
                    .FirstOrDefault();

                if (listAttribute != null)
                {
                    AddListValidator(validationRange,
                                     literalListValidatorFormulae[column],
                                     listAttribute.AllowBlank,
                                     listAttribute.ShowInputMessage,
                                     listAttribute.ShowErrorMessage);
                    continue;
                }

                ExcelListFormulaValidatorAttribute listFormulaAttribute = column
                    .GetCustomAttributes(typeof(ExcelListFormulaValidatorAttribute), false)
                    .Cast<ExcelListFormulaValidatorAttribute>()
                    .FirstOrDefault();

                if (listFormulaAttribute != null)
                {
                    AddListValidator(validationRange,
                                     listFormulaAttribute.Formula,
                                     listFormulaAttribute.AllowBlank,
                                     listFormulaAttribute.ShowInputMessage,
                                     listFormulaAttribute.ShowErrorMessage);
                    continue;
                }

                ExcelIntegerValidatorAttribute integerAttribute = column
                    .GetCustomAttributes(typeof(ExcelIntegerValidatorAttribute), false)
                    .Cast<ExcelIntegerValidatorAttribute>()
                    .FirstOrDefault();

                if (integerAttribute != null)
                {
                    AttachIntegerValidatorToSheet(integerAttribute, validationRange);
                    continue;
                }

                ExcelDecimalValidatorAttribute decimalAttribute = column
                    .GetCustomAttributes(typeof(ExcelDecimalValidatorAttribute), false)
                    .Cast<ExcelDecimalValidatorAttribute>()
                    .FirstOrDefault();

                if (decimalAttribute != null)
                {
                    AttachDecimalValidatorToSheet(decimalAttribute, validationRange);
                    continue;
                }
            }
        }

        private void AttachIntegerValidatorToSheet(ExcelIntegerValidatorAttribute integerAttribute, CellRange validationRange)
        {
            DataValidationOperatorValues validationOperator = integerAttribute.ValidationType.ToDataValidationOperatorValues();
            switch (integerAttribute.FirstOperand)
            {
                case int firstOperand:
                    int? intSecondOperand = validationOperator.RequiresSecondOperand() ? (int)integerAttribute.SecondOperand : (int?)null;
                    AddIntegerValidator(validationRange, firstOperand, validationOperator, integerAttribute.AllowBlank,
                                        integerAttribute.ShowInputMessage, integerAttribute.ShowErrorMessage, intSecondOperand);
                    break;
                case uint firstOperand:
                    uint? uintSecondOperand = validationOperator.RequiresSecondOperand() ? (uint)integerAttribute.SecondOperand : (uint?)null;
                    AddIntegerValidator(validationRange, firstOperand, validationOperator, integerAttribute.AllowBlank,
                                        integerAttribute.ShowInputMessage, integerAttribute.ShowErrorMessage, uintSecondOperand);
                    break;
                case long firstOperand:
                    long? longSecondOperand = validationOperator.RequiresSecondOperand() ? (long)integerAttribute.SecondOperand : (long?)null;
                    AddIntegerValidator(validationRange, firstOperand, validationOperator, integerAttribute.AllowBlank,
                                        integerAttribute.ShowInputMessage, integerAttribute.ShowErrorMessage, longSecondOperand);
                    break;
                case ulong firstOperand:
                    ulong? ulongSecondOperand = validationOperator.RequiresSecondOperand() ? (ulong)integerAttribute.SecondOperand : (ulong?)null;
                    AddIntegerValidator(validationRange, firstOperand, validationOperator, integerAttribute.AllowBlank,
                                        integerAttribute.ShowInputMessage, integerAttribute.ShowErrorMessage, ulongSecondOperand);
                    break;
                default:
                    throw new NotSupportedException("Unsupported integer validator operand type.");
            }
        }

        private void AttachDecimalValidatorToSheet(ExcelDecimalValidatorAttribute decimalAttribute, CellRange validationRange)
        {
            DataValidationOperatorValues validationOperator = decimalAttribute.ValidationType.ToDataValidationOperatorValues();
            switch (decimalAttribute.FirstOperand)
            {
                case decimal firstOperand:
                    decimal? decimalSecondOperand = validationOperator.RequiresSecondOperand() ? (decimal)decimalAttribute.SecondOperand : (decimal?)null;
                    AddDecimalValidator(validationRange, firstOperand, validationOperator, decimalAttribute.AllowBlank,
                                        decimalAttribute.ShowInputMessage, decimalAttribute.ShowErrorMessage, decimalSecondOperand);
                    break;
                case float firstOperand:
                    float? floatSecondOperand = validationOperator.RequiresSecondOperand() ? (float)decimalAttribute.SecondOperand : (float?)null;
                    AddDecimalValidator(validationRange, firstOperand, validationOperator, decimalAttribute.AllowBlank,
                                        decimalAttribute.ShowInputMessage, decimalAttribute.ShowErrorMessage, floatSecondOperand);
                    break;
                case double firstOperand:
                    double? doubleSecondOperand = validationOperator.RequiresSecondOperand() ? (double)decimalAttribute.SecondOperand : (double?)null;
                    AddDecimalValidator(validationRange, firstOperand, validationOperator, decimalAttribute.AllowBlank,
                                        decimalAttribute.ShowInputMessage, decimalAttribute.ShowErrorMessage, doubleSecondOperand);
                    break;
                default:
                    throw new NotSupportedException("Unsupported decimal validator operand type.");
            }
        }

        private string CreateLiteralListValidationSheet(string hiddenSheetName, string[] values)
        {
            CreateAndOpenSheet(hiddenSheetName, null, SheetStateValues.VeryHidden);

            int valueCount = values.Length > 0 ? values.Length : 1;
            for (int index = 0; index < valueCount; index++)
            {
                BeginRow();
                string value = values.Length > 0 ? values[index] ?? string.Empty : string.Empty;
                WriteTextCell(value);
                EndRow();
            }

            CloseSheet();

            return "'" + hiddenSheetName + "'!$A$1:$A$" + valueCount.ToString(CultureInfo.InvariantCulture);
        }
    }
}
