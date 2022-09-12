using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BigExcelCreator.Styles
{
    public class StyleList
    {
        #region props
        private List<Font> Fonts { get; } = new List<Font>();

        private List<Fill> Fills { get; } = new List<Fill>();

        private List<Border> Borders { get; } = new List<Border>();

        private List<NumberingFormat> NumberingFormats { get; } = new List<NumberingFormat>();

        public IList<StyleElement> Styles { get; } = new List<StyleElement>();
        #endregion

        #region ctor
        public StyleList()
        {
            //Create default style
            Font defaultFont = new(
                        new FontSize { Val = 11 },
                        new Color { Rgb = new HexBinaryValue { Value = "000000" } },
                        new FontName { Val = "Calibri" });
            Fill defaultFill = new(
                        new PatternFill { PatternType = PatternValues.None });
            Fill defaultFillGray125 = new(
                        new PatternFill { PatternType = PatternValues.Gray125 });
            Border defaultBorder = new(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder());
            NumberingFormat defaultNumberingFormat = new() { NumberFormatId = 164, FormatCode = "0,.00;(0,.00)" };
            NewStyle(defaultFont, defaultFill, defaultBorder, defaultNumberingFormat, "DEFAULT");
            /* https://stackoverflow.com/a/42789914/14217380
             * For some reason I cannot seem to find documented, Fill Id 0 will always be None,
             * and Fill Id 1 will always be Gray125. If you want a custom fill,
             * you will need to get to at least Fill Id 2.
             */
            NewStyle(defaultFont, defaultFillGray125, defaultBorder, defaultNumberingFormat, "DEFAULTFillGray125");
        }
        #endregion

        public Stylesheet GetStylesheet()
        {
#if NET35
            return new Stylesheet
            {
                Fonts = new Fonts(Fonts.Select(f  => (OpenXmlElement)f)),
                Fills = new Fills(Fills.Select(f  => (OpenXmlElement)f)),
                Borders = new Borders(Borders.Select(b  => (OpenXmlElement)b)),
                NumberingFormats = new NumberingFormats(NumberingFormats.Select(nf  => (OpenXmlElement)nf)),
                CellFormats = new CellFormats(Styles.Select(x => (OpenXmlElement)x.Style)),
            };
#else
            return new Stylesheet
            {
                Fonts = new Fonts(Fonts),
                Fills = new Fills(Fills),
                Borders = new Borders(Borders),
                NumberingFormats = new NumberingFormats(NumberingFormats),
                CellFormats = new CellFormats(Styles.Select(x => x.Style)),
            };
#endif
        }

        public int GetIndexByName(string name)
        {
            return GetIndexByName(name, out _);
        }

        public int GetIndexByName(string name, out StyleElement styleElement)
        {
            styleElement = Styles.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            return Styles.IndexOf(styleElement);
        }

        public StyleElement NewStyle(Font font, Fill fill, Border border, NumberingFormat numberingFormat, string name)
        {
            return NewStyle(font, fill, border, numberingFormat, null, name);
        }

        public StyleElement NewStyle(Font font, Fill fill, Border border, NumberingFormat numberingFormat, Alignment alignment, string name)
        {
            if (GetIndexByName(name, out StyleElement style) >= 0)
            {
                return style;
            }

            int fontId, fillId, borderId, numberingFormatId = 0;

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

            if (numberingFormat != null)
            {
                if ((numberingFormatId = NumberingFormats.IndexOf(numberingFormat)) < 0)
                {
                    numberingFormatId = NumberingFormats.Count;
                    NumberingFormats.Add(numberingFormat);
                }
            }

            return NewStyle(fontId, fillId, borderId, numberingFormatId, alignment, name);
        }

        public StyleElement NewStyle(int? fontId, int? fillId, int? borderId, int? numberingFormatId, Alignment alignment, string name)
        {
            if (fontId < 0) { throw new ArgumentOutOfRangeException(nameof(fontId), "must be greater than 0"); }
            if (fillId < 0) { throw new ArgumentOutOfRangeException(nameof(fillId), "must be greater than 0"); }
            if (borderId < 0) { throw new ArgumentOutOfRangeException(nameof(borderId), "must be greater than 0"); }
            if (numberingFormatId < 0) { throw new ArgumentOutOfRangeException(nameof(numberingFormatId), "must be greater than 0"); }

            StyleElement styleElement = new(name, fontId, fillId, borderId, numberingFormatId, alignment);

            Styles.Add(styleElement);

            return styleElement;
        }
    }
}
