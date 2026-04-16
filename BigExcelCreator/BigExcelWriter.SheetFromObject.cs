// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.ClassAttributes;
using BigExcelCreator.ClassAttributes.Interfaces;
using BigExcelCreator.Enums;
using BigExcelCreator.Exceptions;
using BigExcelCreator.Ranges;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
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
        /// </list>
        /// <para>The sheet state is set to <see cref="SheetStateValues.Visible"/> by default.</para>
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="data"/> is <c>null</c>.</exception>
        /// <exception cref="SheetAlreadyOpenException">Thrown when a sheet is already open and not closed before opening a new one.</exception>
        /// <exception cref="SheetNameCannotBeEmptyException">Thrown when <paramref name="sheetName"/> is null or empty.</exception>
        /// <exception cref="SheetWithSameNameAlreadyExistsException">Thrown when a sheet with the same name already exists.</exception>
        public void CreateSheetFromObject<T>(IEnumerable<T> data, string sheetName, bool writeHeaderRow = true, bool addAutoFilterOnFirstColumn = false, IList<Column> columns = default, bool useSharedStringsOnTextData = false)
             where T : class => CreateSheetFromObject(data, sheetName, SheetStateValues.Visible, writeHeaderRow, addAutoFilterOnFirstColumn, columns, useSharedStringsOnTextData);

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
        /// </list>
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="data"/> is <c>null</c>.</exception>
        /// <exception cref="SheetAlreadyOpenException">Thrown when a sheet is already open and not closed before opening a new one.</exception>
        /// <exception cref="SheetNameCannotBeEmptyException">Thrown when <paramref name="sheetName"/> is null or empty.</exception>
        /// <exception cref="SheetWithSameNameAlreadyExistsException">Thrown when a sheet with the same name already exists.</exception>
        public void CreateSheetFromObject<T>(IEnumerable<T> data, string sheetName, SheetStateValues sheetState, bool writeHeaderRow = true, bool addAutoFilterOnFirstColumn = false, IList<Column> columns = default, bool useSharedStringsOnTextData = false)
            where T : class
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(data);
#else
            if (data is null) throw new ArgumentNullException(nameof(data));
#endif
            IList<T> list = data as IList<T> ?? [.. data];

            if (columns?.Any() != true) { columns = CreateSpreadsheetColumnsFromObject(typeof(T)); }

            CreateAndOpenSheet(sheetName, columns, sheetState);

            List<PropertyInfo> sortedColumns = [.. GetOrderedColumnProperties(typeof(T))];

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

            CloseSheet();
        }
    }
}
