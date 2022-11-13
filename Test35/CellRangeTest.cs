using BigExcelCreator.Ranges;
using NUnit.Framework;
using System;

namespace Test35
{
    internal class CellRangeTest
    {
        [SetUp]
        public void Setup()
        {
            // Method intentionally left empty.
        }


        [TestCase("A1:c5")]
        [TestCase("A1:c$5")]
        [TestCase("A1:$c5")]
        [TestCase("A1:$c$5")]
        [TestCase("A$1:c5")]
        [TestCase("$A1:c$5")]
        [TestCase("$A$1:$c5")]
        [TestCase("A$1:$c$5")]
        [TestCase("$A$1:$c$5")]
        [TestCase("$A$1:$c5")]
        [TestCase("$A$1:c$5")]
        [TestCase("Sheet!A1:c5")]
        [TestCase("A4355:z315")]
        [TestCase("Aa1:ca5")]
        [TestCase("z1:z5")]
        [TestCase("Aa1:ZZ5")]
        [TestCase("ers241:ouy35")]
        [TestCase("ers:ouy")]
        [TestCase("241:35")]
        [TestCase("vals!$A$1:$A$6")]
        public void Parse(string rangeStr)
        {
            CellRange r = new CellRange(rangeStr);
            Assert.That(r.RangeString, Is.EqualTo(rangeStr).IgnoreCase);
        }


        [TestCase("A1:c5")]
        [TestCase("A1:c$5")]
        [TestCase("A1:$c5")]
        [TestCase("A1:$c$5")]
        [TestCase("A$1:c5")]
        [TestCase("$A1:c$5")]
        [TestCase("$A$1:$c5")]
        [TestCase("A$1:$c$5")]
        [TestCase("$A$1:$c$5")]
        [TestCase("$A$1:$c5")]
        [TestCase("$A$1:c$5")]
        [TestCase("Sheet!A1:c5")]
        [TestCase("Sheet!A$1:c5")]
        [TestCase("Sheet!$A1:c$5")]
        [TestCase("A4355:z315")]
        [TestCase("Aa1:ca5")]
        [TestCase("z1:z5")]
        [TestCase("Aa1:ZZ5")]
        [TestCase("ers241:ouy35")]
        [TestCase("ers:ouy")]
        [TestCase("241:35")]
        [TestCase("vals!$A$1:$A$6")]
        public void Equivalence(string rangeStr)
        {
            CellRange parsedRange = new CellRange(rangeStr);

            CellRange createdRange = new CellRange(parsedRange.StartingColumn, parsedRange.StartingColumnIsFixed,
                                                   parsedRange.StartingRow, parsedRange.StartingRowIsFixed,
                                                   parsedRange.EndingColumn, parsedRange.EndingColumnIsFixed,
                                                   parsedRange.EndingRow, parsedRange.EndingRowIsFixed,
                                                   parsedRange.Sheetname);

            Assert.Multiple(() =>
            {
                Assert.That(parsedRange.RangeString, Is.EqualTo(rangeStr).IgnoreCase);
                Assert.That(createdRange.RangeString, Is.EqualTo(rangeStr).IgnoreCase);
                Assert.That(createdRange, Is.EqualTo(parsedRange));
                Assert.That(createdRange.RangeString, Is.EqualTo(parsedRange.RangeString));
                Assert.That(createdRange.RangeStringNoSheetName, Is.EqualTo(parsedRange.RangeStringNoSheetName));
                Assert.That(createdRange.GetHashCode(), Is.EqualTo(parsedRange.GetHashCode()));
            });
        }


        [TestCase("1:c5")]
        [TestCase("!A1:c5")]
        [TestCase("A:z315")]
        [TestCase("Aa1:ca")]
        [TestCase("z1:5")]
        [TestCase("ers241:ouy35!Sheet")]
        public void Error(string rangeStr)
        {
            Assert.Throws<InvalidRangeException>(() => new CellRange(rangeStr));
        }


        [TestCase("qw123:qw123", "qw123")]
        [TestCase("qw123", "qw123")]
        public void SingleRangeString(string rangeStr, string expectedRange)
        {
            CellRange cellRange = new CellRange(rangeStr);
            CellRange cellRangeExpected = new CellRange(expectedRange);
            Assert.Multiple(() =>
            {
                Assert.That(cellRange.RangeString, Is.EqualTo(expectedRange).IgnoreCase);
                Assert.That(cellRangeExpected.RangeString, Is.EqualTo(expectedRange).IgnoreCase);
                Assert.That(cellRange, Is.EqualTo(cellRangeExpected));
                Assert.That(cellRange.GetHashCode(), Is.EqualTo(cellRangeExpected.GetHashCode()));
            });
        }



        [TestCase("a1", "a2")]
        [TestCase("b1", "a2")]
        [TestCase("a2:b5", "a2:j7")]
        [TestCase("a2:j7", "a3:b5")]
        [TestCase("a2:a2", "a2:b5")]
        public void Order(string a, string b)
        {
            Assert.That(new CellRange(a), Is.LessThan(new CellRange(b)));
        }


        [TestCase("A1:c5", 3, 5)]
        [TestCase("A1:c$5", 3, 5)]
        [TestCase("A1:$c5", 3, 5)]
        [TestCase("A1:$c$5", 3, 5)]
        [TestCase("A$1:c5", 3, 5)]
        [TestCase("$A1:c$5", 3, 5)]
        [TestCase("$A$1:$c5", 3, 5)]
        [TestCase("A$1:$c$5", 3, 5)]
        [TestCase("$A$1:$c$5", 3, 5)]
        [TestCase("$A$1:$c5", 3, 5)]
        [TestCase("$A$1:c$5", 3, 5)]
        [TestCase("Hoja!A1:c5", 3, 5)]
        [TestCase("A43:z31", 26, 13)]
        [TestCase("Aa1:ca5", 53, 5)]
        [TestCase("z1:z5", 1, 5)]
        [TestCase("Aa1:ZZ5", 676, 5)]
        public void Size(string rangeStr, int expectedWidth, int expectedHeight)
        {
            CellRange range = new CellRange(rangeStr);
            Assert.Multiple(() =>
            {
                Assert.That(range.Width, Is.EqualTo(expectedWidth));
                Assert.That(range.Height, Is.EqualTo(expectedHeight));
            });
        }

        [TestCase(0, 0, 0, 0)]
        [TestCase(0, 0, 0, 1)]
        [TestCase(0, 0, 1, 0)]
        [TestCase(0, 1, 0, 0)]
        [TestCase(1, 0, 0, 0)]
        [TestCase(1, 1, 1, 0)]
        [TestCase(1, 1, 0, 1)]
        [TestCase(1, 0, 1, 1)]
        [TestCase(0, 1, 1, 1)]
        public void InvalidArgument(int startC, int startR, int endC, int endR)
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new CellRange(startC, startR, endC, endR, ""));
        }

        [TestCase(1, 1, 1, null)]
        [TestCase(1, 1, null, 1)]
        [TestCase(1, 1, null, null)]
        [TestCase(1, null, 1, 1)]
        [TestCase(null, 1, 1, 1)]
        [TestCase(null, null, 1, 1)]
        public void InvalidRange(int? startC, int? startR, int? endC, int? endR)
        {
            Assert.Throws<InvalidRangeException>(() => new CellRange(startC, startR, endC, endR, ""));
        }

        [Test]
        public void RangeOK()
        {
            Assert.DoesNotThrow(() => new CellRange(1, 1, 1, 1, ""));
        }


        [TestCase("a:a", "a:a")]
        [TestCase("a:a", "1:1")]
        [TestCase("a2:a5", "a1:a3")]
        [TestCase("a2:a5", "a1:a13")]
        [TestCase("a2:a5", "a3:a4")]
        [TestCase("a2:d2", "a1:a3")]
        public void OverlappingRanges(string r1, string r2)
        {
            CellRange r1r = new CellRange(r1);
            CellRange r2r = new CellRange(r2);
            Assert.Multiple(() =>
            {
                Assert.That(r1r.RangeOverlaps(r2r), Is.True);
                Assert.That(r2r.RangeOverlaps(r1r), Is.True);
            });
        }

        [TestCase("A:A", "B:B")]
        [TestCase("A:A", "B1:D7")]
        [TestCase("A1:D5", "B10:D70")]
        public void NonOverlappingRanges(string r1, string r2)
        {
            CellRange r1r = new CellRange(r1);
            CellRange r2r = new CellRange(r2);
            Assert.Multiple(() =>
            {
                Assert.That(r1r.RangeOverlaps(r2r), Is.False);
                Assert.That(r2r.RangeOverlaps(r1r), Is.False);
            });
        }
    }
}
