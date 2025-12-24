// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.ClassAttributes;
using BigExcelCreator.Enums;
using BigExcelCreator.Exceptions;
using BigExcelCreator.Extensions;
using BigExcelCreator.Ranges;
using DocumentFormat.OpenXml;
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
        private DataValidation AddValidatorCommon(CellRange range,
                                        DataValidationValues dataValidationValue,
                                        DataValidationOperatorValues validationType,
                                        bool allowBlank = true,
                                        bool showInputMessage = true,
                                        bool showErrorMessage = true)
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(range);
#else
            if (range == null) { throw new ArgumentNullException(nameof(range)); }
#endif
            if (!sheetOpen) { throw new NoOpenSheetException(string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoWordsConcatenation, "Validators", ConstantsAndTexts.NeedToBePlacedOnSSheet)); }

            sheetDataValidations ??= new DataValidations();
            DataValidation dataValidation = new()
            {
                Type = dataValidationValue,
                AllowBlank = allowBlank,
                Operator = validationType,
                ShowInputMessage = showInputMessage,
                ShowErrorMessage = showErrorMessage,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = range.RangeString },
            };

            return dataValidation;
        }

        private void AppendNewDataValidation(DataValidation dataValidation, string firstOperand, string secondOperand = null)
        {
            sheetDataValidations ??= new DataValidations();

            Formula1 formula1 = new() { Text = firstOperand };
            dataValidation.Append(formula1);

            if (dataValidation.Operator.Value.RequiresSecondOperand())
            {
                Formula2 formula2 = new() { Text = secondOperand };
                dataValidation.Append(formula2);
            }
            sheetDataValidations.Append(dataValidation);
            sheetDataValidations.Count = (sheetDataValidations.Count ?? 0) + 1;
        }

        private void WriteFilters()
        {
            if (SheetAutoFilter == null) { return; }

            workSheetPartWriter.WriteElement(SheetAutoFilter);

            SheetAutoFilter = null;
        }

        private void WriteNumberCellInternal(string number, int format = 0)
        {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!rowOpen) { throw new NoOpenRowException(ConstantsAndTexts.NoActiveRow); }

            //reset the list of attributes
            List<OpenXmlAttribute> attributes =
            [
                //add the cell reference attribute
                new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
                //styles
                new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
            ];

            //write the cell start element with the type and reference attributes
            workSheetPartWriter.WriteStartElement(new Cell(), attributes);
            //write the cell value
            workSheetPartWriter.WriteElement(new CellValue(number));

            // write the end cell element
            workSheetPartWriter.WriteEndElement();

            columnNum++;
        }

        private void WriteValidations()
        {
            if (sheetDataValidations == null) { return; }

            workSheetPartWriter.WriteStartElement(sheetDataValidations);
            foreach (DataValidation dataValidation in sheetDataValidations.ChildElements.Cast<DataValidation>())
            {
                workSheetPartWriter.WriteStartElement(dataValidation);
                workSheetPartWriter.WriteElement(dataValidation.Formula1);
                if (dataValidation.Formula2 != null) { workSheetPartWriter.WriteElement(dataValidation.Formula2); }
                workSheetPartWriter.WriteEndElement();
            }
            workSheetPartWriter.WriteEndElement();

            sheetDataValidations = null;
        }

        private void WriteConditionalFormatting()
        {
            if (conditionalFormattingList == null || conditionalFormattingList.Count == 0) { return; }

            foreach (ConditionalFormatting conditionalFormatting in conditionalFormattingList)
            {
                workSheetPartWriter.WriteStartElement(conditionalFormatting);
                foreach (ConditionalFormattingRule conditionalFormattingRule in conditionalFormatting.ChildElements.OfType<ConditionalFormattingRule>())
                {
                    workSheetPartWriter.WriteStartElement(conditionalFormattingRule);
                    foreach (OpenXmlElement item in conditionalFormattingRule.ChildElements)
                    {
                        workSheetPartWriter.WriteElement(item);
                    }
                    workSheetPartWriter.WriteEndElement();
                }
                workSheetPartWriter.WriteEndElement();
            }

            conditionalFormattingList.Clear();
        }

        private void WriteMergedCells()
        {
            if (SheetMergedCells == null || SheetMergedCells.Count == 0) { return; }

            workSheetPartWriter.WriteStartElement(new MergeCells());

            foreach (CellRange range in SheetMergedCells)
            {
                workSheetPartWriter.WriteElement(new MergeCell { Reference = range.RangeStringNoSheetName });
            }

            workSheetPartWriter.WriteEndElement();

            SheetMergedCells.Clear();
        }

        private int AddTextToSharedStringsTable(string text)
        {
            int pos = SharedStringsList.IndexOf(text);
            if (pos < 0)
            {
                pos = SharedStringsList.Count;
                SharedStringsList.Add(text);
            }
            return pos;
        }

        private void WriteSharedStringsPart()
        {
            using OpenXmlWriter SharedStringsWriter = OpenXmlWriter.Create(SharedStringTablePart);
            SharedStringsWriter.WriteStartElement(new SharedStringTable());
            foreach (string item in SharedStringsList)
            {
                SharedStringsWriter.WriteStartElement(new SharedStringItem());
                SharedStringsWriter.WriteElement(new Text(item));
                SharedStringsWriter.WriteEndElement();
            }
            SharedStringsWriter.WriteEndElement();
            SharedStringsWriter.Close();
        }

        private void WriteSheetsAndClosePart()
        {
            using OpenXmlWriter workbookPartWriter = OpenXmlWriter.Create(Document.WorkbookPart);
            workbookPartWriter.WriteStartElement(new Workbook());
            workbookPartWriter.WriteStartElement(new Sheets());

            foreach (Sheet sheet in sheets)
            {
                workbookPartWriter.WriteElement(sheet);
            }

            // End Sheets
            workbookPartWriter.WriteEndElement();
            // End Workbook
            workbookPartWriter.WriteEndElement();
            workbookPartWriter.Close();
        }

        private void SetSheetDefaults()
        {
            _showGridLinesInCurrentSheet = _showGridLinesDefault;
            _showRowAndColumnHeadingsInCurrentSheet = _showRowAndColumnHeadingsDefault;
            _printRowAndColumnHeadingsInCurrentSheet = _printRowAndColumnHeadingsDefault;
            _printGridLinesInCurrentSheet = _printGridLinesDefault;
        }

        private void WritePageConfig(Worksheet worksheet)
        {
            if (_showGridLinesInCurrentSheet != _showGridLinesDefault || _showRowAndColumnHeadingsInCurrentSheet != _showRowAndColumnHeadingsDefault)
            {
                SheetView sheetView = new();
                if (_showGridLinesInCurrentSheet != _showGridLinesDefault) { sheetView.ShowGridLines = _showGridLinesInCurrentSheet; }
                if (_showRowAndColumnHeadingsInCurrentSheet != _showRowAndColumnHeadingsDefault) { sheetView.ShowRowColHeaders = _showRowAndColumnHeadingsInCurrentSheet; }
                sheetView.WorkbookViewId = 0;

                worksheet.SheetViews = new(sheetView);
            }
        }

        private void WritePrintOptions()
        {
            if (_printGridLinesInCurrentSheet != _printGridLinesDefault || _printRowAndColumnHeadingsInCurrentSheet != _printRowAndColumnHeadingsDefault)
            {
                PrintOptions printOptions = new();
                if (_printGridLinesInCurrentSheet != _printGridLinesDefault) { printOptions.GridLines = _printGridLinesInCurrentSheet; }
                if (_printRowAndColumnHeadingsInCurrentSheet != _printRowAndColumnHeadingsDefault) { printOptions.Headings = _printRowAndColumnHeadingsInCurrentSheet; }
                workSheetPartWriter.WriteElement(printOptions);
            }
        }

        private static void ThrowIfInvalidSpreadsheetDocumentType(SpreadsheetDocumentType spreadsheetDocumentType)
        {
            SpreadsheetDocumentType[] validSpreadsheetDocumentTypes =
            [
                SpreadsheetDocumentType.Workbook,
                SpreadsheetDocumentType.Template,
                SpreadsheetDocumentType.MacroEnabledWorkbook,
                SpreadsheetDocumentType.MacroEnabledTemplate,
            ];
            if (!validSpreadsheetDocumentTypes.Contains(spreadsheetDocumentType))
            {
                throw new UnsupportedSpreadsheetDocumentTypeException(string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.InvalidSpreadsheetDocumentType, spreadsheetDocumentType));
            }
        }

        private static List<Column> CreateColumnsFromObject(Type type)
        {
            List<Column> columns = [];

            var dtoCols = GetColumnsOrdered(type);
            foreach (var dtoCol in dtoCols)
            {
                bool hidden = Attribute.IsDefined(dtoCol, typeof(ExcelColumnHiddenAttribute));
                int? width = dtoCol.GetCustomAttributes(typeof(ExcelColumnWidthAttribute), false)
                  .Cast<ExcelColumnWidthAttribute>()
                  .FirstOrDefault()?
                  .Width;
                bool customWidth = width != null;
                Column column = new() { CustomWidth = customWidth, Hidden = hidden };
                if (customWidth)
                {
                    column.Width = width;
                }

                columns.Add(column);
            }

            return columns;
        }

        private static IOrderedEnumerable<PropertyInfo> GetColumnsOrdered(Type type)
        {
            return type.GetProperties()
                    .Where(x => !Attribute.IsDefined(x, typeof(ExcelIgnoreAttribute)))
                    .OrderBy(x => x.GetCustomAttributes(typeof(ExcelColumnOrderAttribute), false)
                            .Cast<ExcelColumnOrderAttribute>()
                        .FirstOrDefault()?
                        .Order ?? int.MaxValue);
        }

        private void WriteCellFromData(object cellData, CellDataType cellType, int format)
        {
            switch (cellType)
            {
                case CellDataType.Number:
                    if (cellData != null)
                    {
                        WriteNumberCell(cellData, format);
                    }
                    else
                    {
                        WriteTextCell(string.Empty, format);
                    }
                    break;
                case CellDataType.Formula:
                    string cellDataFormula = cellData?.ToString();
                    if (!cellDataFormula.IsNullOrWhiteSpace())
                    {
                        WriteFormulaCell(cellDataFormula, format);
                    }
                    else
                    {
                        WriteTextCell(string.Empty, format);
                    }
                    break;
                case CellDataType.Text:
                    string cellDataString = cellData?.ToString() ?? "";
                    WriteTextCell(cellDataString, format);
                    break;
            }
        }

        private void writeHeaderRowFromData(IOrderedEnumerable<PropertyInfo> sortedColumns, ExcelHeaderStyleFormatAttribute headerFormat)
        {
            BeginRow();
            foreach (var column in sortedColumns)
            {
                string columnName = column.GetCustomAttributes(typeof(ExcelColumnNameAttribute), false).Cast<ExcelColumnNameAttribute>().FirstOrDefault()?.Name ?? column.Name;
                int format = 0;

                ExcelStyleFormatAttribute columnFormat = column.GetCustomAttributes(typeof(ExcelStyleFormatAttribute), false).Cast<ExcelStyleFormatAttribute>().FirstOrDefault();

                if (headerFormat != null)
                {
                    format = headerFormat.Format;
                }

                if (columnFormat != null
                    && (columnFormat.UseStyleInHeader == StylingPriority.Data
                        || headerFormat == null
                    ))
                {
                    format = columnFormat.Format;
                }

                WriteTextCell(columnName, format);
            }
            EndRow();
        }
    }
}
