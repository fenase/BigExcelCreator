using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BigExcelCreator.Styles
{
    public class StyleList
    {

        public List<Font> Fonts { get; set; }

        public List<Fill> Fills { get; set; }

        public List<Border> Borders { get; set; }

        public List<NumberingFormat> NumberingFormats { get; set; }

        public List<StyleElement> Styles { get; set; }



        public StyleList()
        {
            Fonts = new List<Font>();
            Fills = new List<Fill>();
            Borders = new List<Border>();
            NumberingFormats = new List<NumberingFormat>();
            Styles = new List<StyleElement>();
        }

        public Stylesheet GetStylesheet()
        {
            return new Stylesheet
            {
                Fonts = new Fonts(Fonts),
                Fills = new Fills(Fills),
                Borders = new Borders(Borders),
                NumberingFormats = new NumberingFormats(NumberingFormats),
                CellFormats = new CellFormats(Styles.Select(x => x.Style)),
            };
        }


        public int GetIndexByName(string name, out StyleElement styleElement)
        {

            styleElement = Styles.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            return Styles.IndexOf(styleElement);
        }



        public StyleElement NewStyle(Font font, Fill fill, Border border, NumberingFormat numberingFormat, string name)
        {
            if(GetIndexByName(name, out StyleElement style) >= 0)
            {
                return style;
            }


            int fontId, fillId, borderId, numberingFormatId;
            
            if ((fontId = Fonts.IndexOf(font)) < 0)
            {
                fontId = Fonts.Count;
                Fonts.Add(font);
            }
            if ((fillId = Fills.IndexOf(fill)) < 0)
            {
                fillId = Fills.Count;
                Fills.Add(fill);
            }
            if ((borderId = Borders.IndexOf(border)) < 0)
            {
                borderId = Borders.Count;
                Borders.Add(border);
            }
            if ((numberingFormatId = NumberingFormats.IndexOf(numberingFormat)) < 0)
            {
                numberingFormatId = NumberingFormats.Count;
                NumberingFormats.Add(numberingFormat);
            }


            StyleElement styleElement = new StyleElement
            {
                Name = name,
                Style = new CellFormat
                {
                    FontId = (uint)fontId,
                    FillId = (uint)fillId,
                    BorderId = (uint)borderId,
                    NumberFormatId = (uint)numberingFormatId,
                }
            };


            Styles.Add(styleElement);

            return styleElement;
        }
    }
}
