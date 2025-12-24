// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: Validator Validators Autofilter stylesheet finalizer inline unhiding gridlines rownum

using BigExcelCreator.Exceptions;
using BigExcelCreator.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace BigExcelCreator
{
    public partial class BigExcelWriter : IDisposable
    {
        /// <summary>
        /// Writes a text cell to the currently open row in the sheet.
        /// </summary>
        /// <remarks>
        /// In order to use this method, a style list must be provided to the <see cref="BigExcelWriter"/> instance.
        /// </remarks>
        /// <param name="text">The text to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <param name="useSharedStrings">Indicates whether to write the value to the shared strings table. This might help reduce the output file size when the same text is shared multiple times among sheets. Default is false.</param>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        public void WriteTextCell(string text, string styleName, bool useSharedStrings = false)
            => WriteTextCell(text, GetFormatFromStyleName(styleName), useSharedStrings);

        /// <summary>
        /// Writes a text cell to the currently open row in the sheet.
        /// </summary>
        /// <param name="text">The text to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <param name="useSharedStrings">Indicates whether to write the value to the shared strings table. This might help reduce the output file size when the same text is shared multiple times among sheets. Default is false.</param>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        /// <exception cref="ArgumentOutOfRangeException">When <paramref name="format"/> is less than 0</exception>
        public void WriteTextCell(string text, int format = 0, bool useSharedStrings = false)
        {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!rowOpen) { throw new NoOpenRowException(ConstantsAndTexts.NoActiveRow); }

            if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(text)))
            {
                List<OpenXmlAttribute> attributes;
                if (useSharedStrings)
                {
                    string ssPos = AddTextToSharedStringsTable(text).ToString(CultureInfo.InvariantCulture);
                    attributes =
                    [
                        new OpenXmlAttribute("t", null, "s"),
                        new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
                        //styles
                        new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                    ];
                    //write the cell start element with the type and reference attributes
                    workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                    //write the cell value
                    workSheetPartWriter.WriteElement(new CellValue(ssPos));
                }
                else
                {
                    //reset the list of attributes
                    attributes =
                    [
                        // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                        new OpenXmlAttribute("t", null, "str"),
                        //add the cell reference attribute
                        new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture, ConstantsAndTexts.TwoParameterConcatenation, Helpers.GetColumnName(columnNum), lastRowWritten)),
                        //styles
                        new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                    ];
                    //write the cell start element with the type and reference attributes
                    workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                    //write the cell value
                    workSheetPartWriter.WriteElement(new CellValue(text));
                }

                // write the end cell element
                workSheetPartWriter.WriteEndElement();
            }
            columnNum++;
        }

        /// <summary>
        /// Writes a row of text cells to the currently open sheet.
        /// </summary>
        /// <param name="texts">The collection of text strings to write in the row.</param>
        /// <param name="format">The format index to apply to each cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <param name="useSharedStrings">Indicates whether to write the value to the shared strings table. This might help reduce the output file size when the same text is shared multiple times among sheets. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the texts collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        public void WriteTextRow(IEnumerable<string> texts, int format = 0, bool hidden = false, bool useSharedStrings = false)
        {
            BeginRow(hidden);
            foreach (string text in texts ?? throw new ArgumentNullException(nameof(texts)))
            {
                WriteTextCell(text, format, useSharedStrings);
            }
            EndRow();
        }

        /// <summary>
        /// Writes a row of text cells to the currently open sheet.
        /// </summary>
        /// <param name="texts">The collection of text strings to write in the row.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <param name="hidden">Indicates whether the row should be hidden. Default is false.</param>
        /// <param name="useSharedStrings">Indicates whether to write the value to the shared strings table. This might help reduce the output file size when the same text is shared multiple times among sheets. Default is false.</param>
        /// <exception cref="ArgumentNullException">Thrown when the texts collection is null.</exception>
        /// <exception cref="NoOpenSheetException">Thrown when there is no open sheet to write a row to.</exception>
        /// <exception cref="RowAlreadyOpenException">Thrown when a row is already open. Use EndRow to close it.</exception>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        public void WriteTextRow(IEnumerable<string> texts, string styleName, bool hidden = false, bool useSharedStrings = false)
            => WriteTextRow(texts, GetFormatFromStyleName(styleName), hidden, useSharedStrings);
    }
}
