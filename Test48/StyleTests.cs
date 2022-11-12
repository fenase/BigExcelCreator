using BigExcelCreator.Styles;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using System;

namespace Test48
{
    public class StyleTests
    {
        readonly Font[] font = new Font[10];

        readonly Fill[] fill = new Fill[10];

        readonly Border[] border = new Border[10];

        readonly NumberingFormat[] numberingFormat = new NumberingFormat[10];

        readonly Alignment[] alignment = new Alignment[10];


        [SetUp]
        public void Setup()
        {
            font[0] = new Font(new Bold(),
                        new FontSize { Val = 11 },
                        new Color { Rgb = new HexBinaryValue { Value = "000000" } },
                        new FontName { Val = "Calibri" });

            fill[0] = new Fill(new[]{
                        new PatternFill { PatternType = PatternValues.Gray125 } });
            fill[1] = new Fill(new[]{
                        new PatternFill { PatternType = PatternValues.DarkDown } });

            fill[2] = new Fill(new[]{
                        new PatternFill(new[]{
                            new BackgroundColor { Rgb = new HexBinaryValue { Value = "00FF00" } } })
                        { PatternType = PatternValues.Solid } });

            border[0] = new Border(
                            new LeftBorder(new[]{
                                new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } } }
                            )
                            { Style = BorderStyleValues.Thin },
                            new RightBorder(new[]{
                                new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } } }
                            )
                            { Style = BorderStyleValues.Thin },
                            new TopBorder(new[]{
                                new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } } }
                            )
                            { Style = BorderStyleValues.Thin },
                            new BottomBorder(new[]{
                                new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } } }
                            )
                            { Style = BorderStyleValues.Thin },
                        new DiagonalBorder());

            numberingFormat[0] = new NumberingFormat { NumberFormatId = 164, FormatCode = "0,.00;(0,.00)" };

            alignment[0] = new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top };
            alignment[1] = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
        }

        [Test]
        public void RepeatedStyles()
        {
            var list = new StyleList();

            const string name = "nombre";

            _ = list.NewStyle(font[0], fill[0], border[0], numberingFormat[0], name);
            _ = list.NewStyle(font[0], fill[0], border[0], numberingFormat[0], name);



            Assert.That(list.Styles, Has.Count.EqualTo(3));
        }


        [Test]
        public void EqualStyles()
        {
            var list = new StyleList();

            const string name = "nombre";
            const string name2 = "nombre2";
            const string name3 = "nombre3";

            var style1 = list.NewStyle(font[0], fill[0], border[0], numberingFormat[0], alignment[0], name);
            var style2 = list.NewStyle(font[0], fill[0], border[0], numberingFormat[0], alignment[0], name2);

            var diffstyle = list.NewDifferentialStyle(name3, fill: fill[2]);

            Assert.Multiple(() =>
            {
                Assert.That(list.Styles, Has.Count.EqualTo(4));
                Assert.That(list.differentialStyleElements, Has.Count.EqualTo(1));

                Assert.That(list.Styles[2], Is.EqualTo(style1));
                Assert.That(list.Styles[3], Is.EqualTo(style2));

                Assert.That(list.differentialStyleElements[0].Fill, Is.EqualTo(fill[2]));
                Assert.That(list.differentialStyleElements[0].Font, Is.Null);
                Assert.That(list.differentialStyleElements[0].Name, Is.EqualTo(name3));
                Assert.That(list.differentialStyleElements[0].Alignment, Is.Null);
                Assert.That(list.differentialStyleElements[0].NumberingFormat, Is.Null);
                Assert.That(list.differentialStyleElements[0].Border, Is.Null);

                Assert.That(list.differentialStyleElements[0], Is.EqualTo(diffstyle));
            });


            var index1 = list.GetIndexByName(name, out StyleElement styleElement1);
            var index2 = list.GetIndexByName(name2, out StyleElement styleElement2);
            var index1b = list.GetIndexByName(name);
            var index2b = list.GetIndexByName(name2);
            var indexDiff = list.GetIndexDifferentialByName(name3, out DifferentialStyleElement differentialStyleElement);
            var indexDiffb = list.GetIndexDifferentialByName(name3);

            Assert.Multiple(() =>
            {
                Assert.That(index1, Is.EqualTo(2));
                Assert.That(index2, Is.EqualTo(3));
                Assert.That(index1b, Is.EqualTo(index1));
                Assert.That(index2b, Is.EqualTo(index2));
                Assert.That(indexDiffb, Is.EqualTo(indexDiff));
                Assert.That(styleElement1, Is.EqualTo(style1));
                Assert.That(styleElement1.Style, Is.EqualTo(style1.Style));
                Assert.That(styleElement2, Is.EqualTo(style2));
                Assert.That(styleElement2.Style, Is.EqualTo(style2.Style));
                Assert.That(differentialStyleElement, Is.EqualTo(diffstyle));
                Assert.That(differentialStyleElement.DifferentialFormat, Is.EqualTo(diffstyle.DifferentialFormat));
                Assert.That(style1, Is.Not.EqualTo(style2));
                Assert.That(style1.Style, Is.EqualTo(style2.Style));
            });
        }


        [Test]
        public void DifferentStyles()
        {
            var list = new StyleList();

            const string name = "nombre";
            const string name2 = "nombre2";

            var style1 = list.NewStyle(font[0], fill[0], border[0], numberingFormat[0], alignment[0], name);
            var style2 = list.NewStyle(font[0], fill[1], border[0], numberingFormat[0], alignment[0], name2);



            Assert.That(list.Styles, Has.Count.EqualTo(4));

            Assert.Multiple(() =>
            {
                Assert.That(list.Styles[2], Is.EqualTo(style1));
                Assert.That(list.Styles[3], Is.EqualTo(style2));
            });

            var index1 = list.GetIndexByName(name, out StyleElement styleElement1);
            var index2 = list.GetIndexByName(name2, out StyleElement styleElement2);


            Assert.Multiple(() =>
            {
                Assert.That(index1, Is.EqualTo(2));
                Assert.That(index2, Is.EqualTo(3));
                Assert.That(styleElement1, Is.EqualTo(style1));
                Assert.That(styleElement1.Style, Is.EqualTo(style1.Style));
                Assert.That(styleElement2, Is.EqualTo(style2));
                Assert.That(styleElement2.Style, Is.EqualTo(style2.Style));
                Assert.That(style1, Is.Not.EqualTo(style2));
                Assert.That(style1.Style, Is.EqualTo(style2.Style));
            });
        }


        [Test]
        public void SameStylesheet()
        {
            var list1 = new StyleList();
            var list2 = new StyleList();

            const string name = "nombre";
            const string name2 = "nombre2";

            list1.NewStyle(font[0], fill[0], border[0], numberingFormat[0], alignment[0], name);
            list1.NewStyle(font[0], fill[1], border[0], numberingFormat[0], alignment[0], name2);
            list2.NewStyle(font[0], fill[0], border[0], numberingFormat[0], alignment[0], name2);
            list2.NewStyle(font[0], fill[1], border[0], numberingFormat[0], alignment[0], name); //names should not influence the final style sheet
            list1.NewDifferentialStyle("q", fill: fill[2]);
            list2.NewDifferentialStyle("w", fill: fill[2]);

            Assert.That(list1.GetStylesheet(), Is.EqualTo(list2.GetStylesheet()));
        }


        [Test]
        public void SameEmptyStylesheet()
        {
            var list1 = new StyleList();
            var list2 = new StyleList();

            const string name = "nombre";
            const string name2 = "nombre2";

            list1.NewStyle(null, null, null, null, name);
            list1.NewStyle(null, null, null, null, name2);
            list2.NewStyle(null, null, null, null, name2);
            list2.NewStyle(null, null, null, null, name); //names should not influence the final style sheet

            Assert.That(list1.GetStylesheet(), Is.EqualTo(list2.GetStylesheet()));
        }


        [TestCase(-1, -1, -1, -1)]
        [TestCase(0, -1, -1, -1)]
        [TestCase(-1, 0, -1, -1)]
        [TestCase(0, 0, -1, -1)]
        [TestCase(-1, -1, 0, -1)]
        [TestCase(0, -1, 0, -1)]
        [TestCase(-1, 0, 0, -1)]
        [TestCase(0, 0, 0, -1)]
        [TestCase(-1, -1, -1, 0)]
        [TestCase(0, -1, -1, 0)]
        [TestCase(-1, 0, -1, 0)]
        [TestCase(0, 0, -1, 0)]
        [TestCase(-1, -1, 0, 0)]
        [TestCase(0, -1, 0, 0)]
        [TestCase(-1, 0, 0, 0)]
        public void NewStyleError(int? fontId, int? fillId, int? borderId, int? numberingFormatId)
        {
            var list = new StyleList();
            Assert.Multiple(() =>
            {
                Assert.Throws<ArgumentOutOfRangeException>(() => list.NewStyle(fontId, fillId, borderId, numberingFormatId, null, "a"));
                Assert.Throws<ArgumentOutOfRangeException>(() => list.NewStyle(fontId, fillId, borderId, numberingFormatId, new Alignment() { Horizontal = HorizontalAlignmentValues.Center }, "a"));
            });
        }


        [TestCase(0, 0, 0, 0)]
        public void NewStyleOK(int? fontId, int? fillId, int? borderId, int? numberingFormatId)
        {
            var list = new StyleList();
            Assert.Multiple(() =>
            {
                Assert.DoesNotThrow(() => list.NewStyle(fontId, fillId, borderId, numberingFormatId, null, "a"));
                Assert.DoesNotThrow(() => list.NewStyle(fontId, fillId, borderId, numberingFormatId, new Alignment() { Horizontal = HorizontalAlignmentValues.Center }, "a"));
            });
        }


        [Test]
        public void NewStylesListIsNotEmpty()
        {
            var list = new StyleList();
            Assert.That(list.Styles, Has.Count.GreaterThan(0));

            var style = list.NewStyle(new Font(), new Fill(), new Border(), new NumberingFormat(), "");
            Assert.Multiple(() =>
            {
                Assert.That(style.FontIndex, Is.GreaterThan(0));
                Assert.That(style.FillIndex, Is.GreaterThan(0));
                Assert.That(style.BorderIndex, Is.GreaterThan(0));
                Assert.That(style.NumberFormatIndex, Is.GreaterThan(0));
            });
        }
    }
}
