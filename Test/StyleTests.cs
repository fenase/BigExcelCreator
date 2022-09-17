using BigExcelCreator.Styles;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Test
{
    public class StyleTests
    {
        readonly Font[] font = new Font[10];

        readonly Fill[] fill = new Fill[10];

        readonly Border[] border = new Border[10];

        readonly NumberingFormat[] numberingFormat = new NumberingFormat[10];




        [SetUp]
        public void Setup()
        {
            font[0] = new Font(new Bold(),
                        new FontSize { Val = 11 },
                        new Color { Rgb = new HexBinaryValue { Value = "000000" } },
                        new FontName { Val = "Calibri" });

            fill[0] = new Fill(
                        new PatternFill { PatternType = PatternValues.Gray125 });
            fill[1] = new Fill(
                        new PatternFill { PatternType = PatternValues.DarkDown });

            border[0] = new Border(
                        new LeftBorder(
                            new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } }
                        )
                        { Style = BorderStyleValues.Thin },
                        new RightBorder(
                            new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } }
                        )
                        { Style = BorderStyleValues.Thin },
                        new TopBorder(
                            new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } }
                        )
                        { Style = BorderStyleValues.Thin },
                        new BottomBorder(
                            new Color { Rgb = new HexBinaryValue { Value = "FFD3D3D3" } }
                        )
                        { Style = BorderStyleValues.Thin },
                        new DiagonalBorder());

            numberingFormat[0] = new NumberingFormat { NumberFormatId = 164, FormatCode = "0,.00;(0,.00)" };
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

            var style1 = list.NewStyle(font[0], fill[0], border[0], numberingFormat[0], name);
            var style2 = list.NewStyle(font[0], fill[0], border[0], numberingFormat[0], name2);



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
        public void DifferentStyles()
        {
            var list = new StyleList();

            const string name = "nombre";
            const string name2 = "nombre2";

            var style1 = list.NewStyle(font[0], fill[0], border[0], numberingFormat[0], name);
            var style2 = list.NewStyle(font[0], fill[1], border[0], numberingFormat[0], name2);



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
    }
}
