using BigExcelCreator;
using NUnit.Framework;

namespace Test35
{
    internal class HelpersTest
    {
        [SetUp]
        public void Setup()
        {
            // Method intentionally left empty.
        }


        [Test]
        public void GetColumnName()
        {
            Assert.Multiple(() =>
            {
                Assert.That(Helpers.GetColumnName(1), Is.EqualTo("A"));
                Assert.That(Helpers.GetColumnName(2), Is.EqualTo("B"));
                Assert.That(Helpers.GetColumnName(26), Is.EqualTo("Z"));
                Assert.That(Helpers.GetColumnName(27), Is.EqualTo("AA"));
            });
        }

        [Test]
        public void GetColumnIndex()
        {
            Assert.Multiple(() =>
            {
                Assert.That(Helpers.GetColumnIndex("A"), Is.EqualTo(1));
                Assert.That(Helpers.GetColumnIndex("B"), Is.EqualTo(2));
                Assert.That(Helpers.GetColumnIndex("Z"), Is.EqualTo(26));
                Assert.That(Helpers.GetColumnIndex("AA"), Is.EqualTo(27));
            });
        }

        [Test]
        public void GetNameAndIndexAllColumn()
        {
            for (int i = 1; i <= 16384; i++)
            {
                Assert.That(Helpers.GetColumnIndex(Helpers.GetColumnName(i)), Is.EqualTo(i));
            }
        }
    }
}
