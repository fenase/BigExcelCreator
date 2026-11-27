// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.Formulas.InvalidConditionalFormatOperands
{
    [TestFixture]
    internal sealed class InvalidConditionalFormatOperandsAnalyzerTests
    {
        [Test]
        public async Task BetweenWithoutSecondValueIsAnError()
        {
            const string source = "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public class Model { [ExcelConditionalFormatCellIs(0, ConditionalFormattingOperator.Between, \"1\")] " +
                "public int Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL2002"), Is.True);
        }

        [Test]
        public async Task BetweenWithBothValuesIsAccepted()
        {
            const string source = "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public class Model { [ExcelConditionalFormatCellIs(0, ConditionalFormattingOperator.Between, \"1\", \"5\")] " +
                "public int Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL2002"), Is.False);
        }
    }
}
