// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Spreadsheet;

namespace BigExcelCreator.Styles
{
    public class StyleElement
    {
        #region props
        public string Name { get; }

        public CellFormat Style { get; }

        public int FontIndex { get; }
        public int FillIndex { get; }
        public int BorderIndex { get; }
        public int NumberFormatIndex { get; }
        #endregion

        #region ctor
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
