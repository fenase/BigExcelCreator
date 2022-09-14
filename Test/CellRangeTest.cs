using BigExcelCreator.Ranges;

namespace Test
{
    internal class CellRangeTest
    {
        [SetUp]
        public void Setup() { }


        [TestCase("A1:c5")]
        [TestCase("Hoja!A1:c5")]
        [TestCase("A4355:z315")]
        [TestCase("Aa1:ca5")]
        [TestCase("z1:z5")]
        [TestCase("Aa1:zz5")]
        [TestCase("ers241:ouy35")]
        public void Parse(string rangeStr)
        {
            CellRange r = new(rangeStr);
            Assert.That(r.RangeString, Is.EqualTo(rangeStr).IgnoreCase);
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
