using BigExcelCreator.Ranges;
using NUnit.Framework;

namespace Test35
{
    internal class CellRangeTest
    {
        [SetUp]
        public void Setup() { }


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
        [TestCase("Hoja!A1:c5")]
        [TestCase("A4355:z315")]
        [TestCase("Aa1:ca5")]
        [TestCase("z1:z5")]
        [TestCase("Aa1:ZZ5")]
        [TestCase("ers241:ouy35")]
        [TestCase("ers:ouy")]
        [TestCase("241:35")]
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
        [TestCase("Hoja!A1:c5")]
        [TestCase("A4355:z315")]
        [TestCase("Aa1:ca5")]
        [TestCase("z1:z5")]
        [TestCase("Aa1:ZZ5")]
        [TestCase("ers241:ouy35")]
        [TestCase("ers:ouy")]
        [TestCase("241:35")]
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
                Assert.That(createdRange.GetHashCode(), Is.EqualTo(parsedRange.GetHashCode()));
            });
        }


        [TestCase("1:c5")]
        [TestCase("!A1:c5")]
        [TestCase("A:z315")]
        [TestCase("Aa1:ca")]
        [TestCase("z1:5")]
        [TestCase("ers241:ouy35!hoja")]
        public void Error(string rangeStr)
        {
            Assert.Throws<InvalidRangeException>(() => new CellRange(rangeStr));
        }
    }
}
