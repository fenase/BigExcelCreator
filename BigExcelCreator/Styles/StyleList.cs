// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: stylesheet stylesheets Calibri

using BigExcelCreator.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BigExcelCreator.Styles
{
    /// <summary>
    /// Manages styles and generates stylesheets
    /// </summary>
    public class StyleList
    {
        #region props
        private List<Font> Fonts { get; } = [];

        private List<Fill> Fills { get; } = [];

        private List<Border> Borders { get; } = [];

        private List<NumberingFormat> NumberingFormats { get; } = [];

        /// <summary>
        /// Main styles
        /// </summary>
        public IList<StyleElement> Styles { get; } = [];

        /// <summary>
        /// Differential styles.
        /// <para>Used in COnditional formatting</para>
        /// </summary>
        public IList<DifferentialStyleElement> DifferentialStyleElements { get; } = [];

        private const uint STARTINGNUMBERFORMAT = 164;
        #endregion

        #region ctor
        /// <summary>
        /// Creates a style list and populates with default styles
        /// </summary>
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
            _ = NewStyle(defaultFont, defaultFill, defaultBorder, defaultNumberingFormat, "DEFAULT");
            /* https://stackoverflow.com/a/42789914/14217380
             * For some reason I cannot seem to find documented, Fill Id 0 will always be None,
             * and Fill Id 1 will always be Gray125. If you want a custom fill,
             * you will need to get to at least Fill Id 2.
             */
            _ = NewStyle(defaultFont, defaultFillGray125, defaultBorder, defaultNumberingFormat, "DEFAULTFillGray125");
        }
        #endregion

        /// <summary>
        /// Generates a <see cref="Stylesheet"/> to include in an Excel document
        /// </summary>
        /// <returns><see cref="Stylesheet"/>: A stylesheet</returns>
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
                DifferentialFormats = new DifferentialFormats(DifferentialStyleElements.Select(x => (OpenXmlElement)x.DifferentialFormat.Clone())),
            };
#else
            return new Stylesheet
            {
                Fonts = new Fonts(Fonts.Select(x => (Font)x.Clone())),
                Fills = new Fills(Fills.Select(x => (Fill)x.Clone())),
                Borders = new Borders(Borders.Select(x => (Border)x.Clone())),
                NumberingFormats = new NumberingFormats(NumberingFormats.Select(x => (NumberingFormat)x.Clone())),
                CellFormats = new CellFormats(Styles.Select(x => (CellFormat)x.Style.Clone())),
                DifferentialFormats = new DifferentialFormats(DifferentialStyleElements.Select(x => (DifferentialFormat)x.DifferentialFormat.Clone())),
            };
#endif
        }

        /// <summary>
        /// Gets the index of a named style
        /// </summary>
        /// <param name="name">The name of the style to look for.</param>
        /// <returns>The index of the named style, or -1 if not found.</returns>
        public int GetIndexByName(string name) => GetIndexByName(name, out _);

        /// <summary>
        /// Gets the index of a named style.
        /// </summary>
        /// <param name="name">The name of the style to look for.</param>
        /// <param name="styleElement">A copy of the found style.</param>
        /// <returns>The index of the named style, or -1 if not found.</returns>
        public int GetIndexByName(string name, out StyleElement styleElement)
        {
            styleElement = Styles.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            return Styles.IndexOf(styleElement);
        }

        /// <summary>
        /// Gets the index of a named differential style.
        /// </summary>
        /// <param name="name">The name of the differential style to look for.</param>
        /// <returns>The index of the named differential style, or -1 if not found.</returns>
        public int GetIndexDifferentialByName(string name) => GetIndexDifferentialByName(name, out _);

        /// <summary>
        /// Gets the index of a named differential style.
        /// </summary>
        /// <param name="name">The name of the differential style to look for.</param>
        /// <param name="differentialStyleElement">A copy of the found differential style.</param>
        /// <returns>The index of the named differential style, or -1 if not found.</returns>
        public int GetIndexDifferentialByName(string name, out DifferentialStyleElement differentialStyleElement)
        {
            differentialStyleElement = DifferentialStyleElements.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            return DifferentialStyleElements.IndexOf(differentialStyleElement);
        }

        /// <summary>
        /// Generates, stores and returns a new style
        /// </summary>
        /// <param name="font"><see cref="Font"/></param>
        /// <param name="fill"><see cref="Fill"/></param>
        /// <param name="border"><see cref="Border"/></param>
        /// <param name="numberingFormat"><see cref="NumberingFormat"/></param>
        /// <param name="name">A unique name to find the inserted style later</param>
        /// <returns>The <see cref="StyleElement"/> generated</returns>
        public StyleElement NewStyle(Font font, Fill fill, Border border, NumberingFormat numberingFormat, string name)
            => NewStyle(font, fill, border, numberingFormat, null, name);

        /// <summary>
        /// Generates, stores and returns a new style
        /// </summary>
        /// <param name="font"><see cref="Font"/></param>
        /// <param name="fill"><see cref="Fill"/></param>
        /// <param name="border"><see cref="Border"/></param>
        /// <param name="numberingFormat"><see cref="NumberingFormat"/></param>
        /// <param name="alignment"><see cref="Alignment"/></param>
        /// <param name="name">A unique name to find the inserted style later</param>
        /// <returns>The <see cref="StyleElement"/> generated</returns>
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

        /// <summary>
        /// Generates, stores and returns a new style.
        /// </summary>
        /// <remarks>
        /// <para>If the inserted indexes don't exist when the stylesheet is generated, the file might fail to open</para>
        /// <para>To avoid such problems, use <see cref="NewStyle(Font, Fill, Border, NumberingFormat, string)"/> or <see cref="NewStyle(Font, Fill, Border, NumberingFormat, Alignment, string)"/> instead</para>
        /// <para>This method should be private, but it's kept public for backwards compatibility reasons.</para>
        /// </remarks>
        /// <param name="fontId">Index of already inserted font</param>
        /// <param name="fillId">Index of already inserted fill</param>
        /// <param name="borderId">Index of already inserted border</param>
        /// <param name="numberingFormatId">Index of already inserted numbering format</param>
        /// <param name="alignment"><see cref="Alignment"/></param>
        /// <param name="name">A unique name to find the inserted style later</param>
        /// <returns>The <see cref="StyleElement"/> generated</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when any of the provided indexes are less than 0</exception>
        public StyleElement NewStyle(int? fontId, int? fillId, int? borderId, int? numberingFormatId, Alignment alignment, string name)
        {
            if (fontId < 0) { throw new ArgumentOutOfRangeException(nameof(fontId), ConstantsAndTexts.MusBeGreaterThan0); }
            if (fillId < 0) { throw new ArgumentOutOfRangeException(nameof(fillId), ConstantsAndTexts.MusBeGreaterThan0); }
            if (borderId < 0) { throw new ArgumentOutOfRangeException(nameof(borderId), ConstantsAndTexts.MusBeGreaterThan0); }
            if (numberingFormatId < 0) { throw new ArgumentOutOfRangeException(nameof(numberingFormatId), ConstantsAndTexts.MusBeGreaterThan0); }

            StyleElement styleElement = new(name, fontId, fillId, borderId, numberingFormatId, alignment);

            Styles.Add(styleElement);

            return styleElement;
        }

        /// <summary>
        /// Generates, stores and returns a new differential style
        /// </summary>
        /// <param name="name">A unique name to find the inserted style later</param>
        /// <param name="font"><see cref="Font"/></param>
        /// <param name="fill"><see cref="Fill"/></param>
        /// <param name="border"><see cref="Border"/></param>
        /// <param name="numberingFormat"><see cref="NumberingFormat"/></param>
        /// <param name="alignment"><see cref="Alignment"/></param>
        /// <returns>The <see cref="DifferentialStyleElement"/> generated</returns>
        public DifferentialStyleElement NewDifferentialStyle(string name, Font font = null, Fill fill = null, Border border = null, NumberingFormat numberingFormat = null, Alignment alignment = null)
        {
            if (GetIndexDifferentialByName(name, out DifferentialStyleElement style) >= 0)
            {
                return style;
            }

            DifferentialStyleElement differentialFormat = null;
            if (font != null)
            {
                // if this is not the first block, uncomment this and remove the next line => differentialFormat ??= new()
                differentialFormat = new() { Font = font };
            }

            if (fill != null)
            {
                differentialFormat ??= new() { Fill = fill };
            }

            if (border != null)
            {
                differentialFormat ??= new() { Border = border };
            }

            if (numberingFormat != null)
            {
                differentialFormat ??= new() { NumberingFormat = numberingFormat };
            }

            if (alignment != null)
            {
                differentialFormat ??= new() { Alignment = alignment };
            }

            if (differentialFormat != null)
            {
                differentialFormat.Name = !name.IsNullOrWhiteSpace() ? name : throw new ArgumentNullException(nameof(name));
                DifferentialStyleElements.Add(differentialFormat);
                return differentialFormat;
            }
            else
            {
                throw new ArgumentNullException("At least one argument should be not null", (Exception)null);
            }
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
                NumberingFormat nf = NumberingFormats.Find(x => x.FormatCode == numberingFormat.FormatCode);
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
