// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Spreadsheet;

namespace BigExcelCreator.Styles
{
    /// <summary>
    /// A style to be converted to an entry of a stylesheet
    /// </summary>
    public class StyleElement
    {
        #region props
        /// <summary>
        /// Given name of a style
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// A <see cref="CellFormat"/> object representing a style
        /// </summary>
        public CellFormat Style { get; }

        /// <summary>
        /// Font index in the font list of <see cref="StyleList"/>
        /// </summary>
        public int FontIndex { get; }

        /// <summary>
        /// Fill index in the fill list of <see cref="StyleList"/>
        /// </summary>
        public int FillIndex { get; }

        /// <summary>
        /// Border index in the border list of <see cref="StyleList"/>
        /// </summary>
        public int BorderIndex { get; }

        /// <summary>
        /// NumberFormat index in the Number format list of <see cref="StyleList"/>
        /// </summary>
        public int NumberFormatIndex { get; }
        #endregion

        #region ctor
        /// <summary>
        /// The constructor for StyleElement
        /// </summary>
        /// <param name="name"></param>
        /// <param name="fontIndex"></param>
        /// <param name="fillIndex"></param>
        /// <param name="borderIndex"></param>
        /// <param name="numberFormatIndex"></param>
        /// <param name="alignment"></param>
        public StyleElement(string name, int? fontIndex, int? fillIndex, int? borderIndex, int? numberFormatIndex, Alignment alignment)
        {
            Name = name;
            FontIndex = fontIndex ?? 0;
            FillIndex = fillIndex ?? 0;
            BorderIndex = borderIndex ?? 0;
            NumberFormatIndex = numberFormatIndex ?? 0;

            Style = new()
            {
                FontId = (uint)FontIndex,
                FillId = (uint)FillIndex,
                BorderId = (uint)BorderIndex,
                NumberFormatId = (uint)NumberFormatIndex,
            };
            if (alignment != null)
            {
                Style.Alignment = (Alignment)alignment.Clone();
            }
        }
        #endregion
    }
}
