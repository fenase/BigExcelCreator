using BigExcelCreator.Enums;
using BigExcelCreator.Extensions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BigExcelCreator.Test
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

        [Test]
        public void GetConditionalFormattingOperatorValuesValue()
        {
            Assert.Multiple(() =>
            {
                Assert.That(ConditionalFormattingOperator.LessThan.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.LessThan));
                Assert.That(ConditionalFormattingOperator.LessThanOrEqual.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.LessThanOrEqual));
                Assert.That(ConditionalFormattingOperator.Equal.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.Equal));
                Assert.That(ConditionalFormattingOperator.NotEqual.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.NotEqual));
                Assert.That(ConditionalFormattingOperator.GreaterThanOrEqual.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.GreaterThanOrEqual));
                Assert.That(ConditionalFormattingOperator.GreaterThan.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.GreaterThan));
                Assert.That(ConditionalFormattingOperator.Between.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.Between));
                Assert.That(ConditionalFormattingOperator.NotBetween.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.NotBetween));
                Assert.That(ConditionalFormattingOperator.ContainsText.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.ContainsText));
                Assert.That(ConditionalFormattingOperator.NotContains.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.NotContains));
                Assert.That(ConditionalFormattingOperator.BeginsWith.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.BeginsWith));
                Assert.That(ConditionalFormattingOperator.EndsWith.Value(), Is.EqualTo(ConditionalFormattingOperatorValues.EndsWith));

                Assert.Throws<ArgumentOutOfRangeException>(() => ((ConditionalFormattingOperator)999).Value());
                Assert.Throws<ArgumentOutOfRangeException>(() => ((ConditionalFormattingOperator)(-1)).Value());
            });
        }
    }
}
