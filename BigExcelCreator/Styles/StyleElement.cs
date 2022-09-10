using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace BigExcelCreator.Styles
{
    public class StyleElement
    {
        public string Name { get; private set; }

        public CellFormat Style { get; private set; }

        public int? FontIndex { get; private set; }
        public int? FillIndex { get; private set; }
        public int? BorderIndex { get; private set; }
        public int? NumberFormatIndex { get; private set; }



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
    }
}
