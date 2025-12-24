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
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(sbyte number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(sbyte number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(byte number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(byte number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(short number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(short number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(ushort number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(ushort number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(int number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(int number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(uint number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(uint number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(long number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(long number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(ulong number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        [CLSCompliant(false)]
        public void WriteNumberCell(ulong number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(float number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(float number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(double number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(double number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(decimal number, int format = 0)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), format);

        /// <summary>
        /// Writes a numerical value to the currently open row in the sheet.
        /// </summary>
        /// <param name="number">The number to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteNumberCell(decimal number, string styleName)
            => WriteNumberCellInternal(number.ToString(CultureInfo.InvariantCulture), GetFormatFromStyleName(styleName));

        private void WriteNumberCell(object number, int format = 0)
        {
            switch (number)
            {
                case sbyte v: WriteNumberCell(v, format); break;
                case byte v: WriteNumberCell(v, format); break;
                case short v: WriteNumberCell(v, format); break;
                case ushort v: WriteNumberCell(v, format); break;
                case int v: WriteNumberCell(v, format); break;
                case uint v: WriteNumberCell(v, format); break;
                case long v: WriteNumberCell(v, format); break;
                case ulong v: WriteNumberCell(v, format); break;
                case float v: WriteNumberCell(v, format); break;
                case double v: WriteNumberCell(v, format); break;
                case decimal v: WriteNumberCell(v, format); break;
                default:
                    throw new ArgumentException(
                        $"Unsupported numeric type '{number.GetType().FullName}'. Supported types include: sbyte, byte, short, ushort, int, uint, long, ulong, float, double, decimal.",
                        nameof(number));
            }
        }

        private void WriteNumberCell(object number, string styleName)
            => WriteNumberCell(number, GetFormatFromStyleName(styleName));

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
    }
}
