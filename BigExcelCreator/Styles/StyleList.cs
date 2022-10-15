// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

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

        private const uint STARTINGNUMBERFORMAT = 164;
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
            NumberingFormat defaultNumberingFormat = new() { NumberFormatId = STARTINGNUMBERFORMAT, FormatCode = "0,.00;(0,.00)" };
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
                Fonts = new Fonts(Fonts.Select(x => (OpenXmlElement)x.Clone())),
                Fills = new Fills(Fills.Select(x => (OpenXmlElement)x.Clone())),
                Borders = new Borders(Borders.Select(x => (OpenXmlElement)x.Clone())),
                NumberingFormats = new NumberingFormats(NumberingFormats.Select(x => (OpenXmlElement)x.Clone())),
                CellFormats = new CellFormats(Styles.Select(x => (OpenXmlElement)x.Style.Clone())),
            };
#else
            return new Stylesheet
            {
                Fonts = new Fonts(Fonts.Select(x => (Font)x.Clone())),
                Fills = new Fills(Fills.Select(x => (Fill)x.Clone())),
                Borders = new Borders(Borders.Select(x => (Border)x.Clone())),
                NumberingFormats = new NumberingFormats(NumberingFormats.Select(x => (NumberingFormat)x.Clone())),
                CellFormats = new CellFormats(Styles.Select(x => (CellFormat)x.Style.Clone())),
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

            int fontId = GetFontId(font);

            int fillId = GetFillId(fill);

            int borderId = GetBorderId(border);

            int numberingFormatId = GetNumberingFormatId(numberingFormat);

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


        private int GetFontId(Font font)
        {
            int fontId;
            if (font != null)
            {
                if ((fontId = Fonts.IndexOf(font)) < 0)
                {
                    fontId = Fonts.Count;
                    Fonts.Add(font);
                }
            }
            else
            {
                fontId = 0;
            }
            return fontId;
        }

        private int GetFillId(Fill fill)
        {
            int fillId;
            if (fill != null)
            {
                if ((fillId = Fills.IndexOf(fill)) < 0)
                {
                    fillId = Fills.Count;
                    Fills.Add(fill);
                }
            }
            else
            {
                fillId = 0;
            }
            return fillId;
        }

        private int GetBorderId(Border border)
        {
            int borderId;
            if (border != null)
            {
                if ((borderId = Borders.IndexOf(border)) < 0)
                {
                    borderId = Borders.Count;
                    Borders.Add(border);
                }
            }
            else
            {
                borderId = 0;
            }
            return borderId;
        }

        private int GetNumberingFormatId(NumberingFormat numberingFormat)
        {
            int numberingFormatId;
            if (numberingFormat != null)
            {
                NumberingFormat nf = NumberingFormats.FirstOrDefault(x => x.FormatCode == numberingFormat.FormatCode);
                if (nf != null)
                {
                    numberingFormatId = (int)(uint)nf.NumberFormatId;
                }
                else
                {
                    numberingFormatId = (int)Math.Max(STARTINGNUMBERFORMAT, (NumberingFormats.Max(x => x.NumberFormatId) ?? 0) + 1);
                    NumberingFormats.Add(new NumberingFormat() { NumberFormatId = (uint)numberingFormatId, FormatCode = numberingFormat.FormatCode });
                }
            }
            else
            {
                numberingFormatId = 0;
            }
            return numberingFormatId;
        }
    }
}
