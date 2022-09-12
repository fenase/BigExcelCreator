using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace BigExcelCreator.Styles
{
    public class StyleElement
    {
        #region props
        public string Name { get; }

        public CellFormat Style { get; }

        public int? FontIndex { get; }
        public int? FillIndex { get; }
        public int? BorderIndex { get; }
        public int? NumberFormatIndex { get; }
        #endregion

        #region ctor
        public StyleElement(string name, int? fontIndex, int? fillIndex, int? borderIndex, int? numberFormatIndex, Alignment alignment)
        {
            Name = name;
            FontIndex = fontIndex;
            FillIndex = fillIndex;
            BorderIndex = borderIndex;
            NumberFormatIndex = numberFormatIndex;

            Style = new()
            {
                FontId = (uint)fontIndex,
                FillId = (uint)fillIndex,
                BorderId = (uint)borderIndex,
                NumberFormatId = 0,
            };
            if (numberFormatIndex != null)
            {
                Style.NumberFormatId = (uint)numberFormatIndex;
            }
            if (alignment != null)
            {
                Style.Alignment = alignment;
            }
        }
        #endregion
    }
}
