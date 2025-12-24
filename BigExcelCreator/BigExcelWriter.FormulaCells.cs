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
        /// Writes a formula cell to the currently open row in the sheet.
        /// </summary>
        /// <param name="formula">The formula to write in the cell.</param>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/></param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="format"/> is less than 0</exception>
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteFormulaCell(string formula, int format = 0)
        {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegative(format);
#else
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
#endif
            if (!rowOpen) { throw new NoOpenRowException(ConstantsAndTexts.NoActiveRow); }

            if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(formula)))
            {
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
                workSheetPartWriter.WriteElement(new CellFormula(formula?.ToUpperInvariant()));

                // write the end cell element
                workSheetPartWriter.WriteEndElement();
            }
            columnNum++;
        }


        /// <summary>
        /// Writes a formula cell to the currently open row in the sheet.
        /// </summary>
        /// <param name="formula">The formula to write in the cell.</param>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <exception cref="StyleListNotAvailableException">Thrown when no style list was provided to the <see cref="BigExcelWriter"/> instance.</exception>
        /// <exception cref="StyleNameMustBeProvidedException">Thrown when <paramref name="styleName"/>is empty.</exception>
        /// <exception cref="StyleNameNotFoundException">Thrown when the provided style name was not found in the style list.</exception>"
        /// <exception cref="NoOpenRowException">Thrown when there is no open row to write the cell to.</exception>
        public void WriteFormulaCell(string formula, string styleName)
            => WriteFormulaCell(formula, GetFormatFromStyleName(styleName));
    }
}
