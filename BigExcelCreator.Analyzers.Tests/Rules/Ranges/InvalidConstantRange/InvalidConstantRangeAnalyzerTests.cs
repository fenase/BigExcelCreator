// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.Ranges.InvalidConstantRange
{
    [TestFixture]
    internal sealed class InvalidConstantRangeAnalyzerTests
    {
        [Test]
        public async Task InvalidConstantSheetNameAndRangeAreErrors()
        {
            const string source = "using BigExcelCreator; using BigExcelCreator.Ranges; public class Model { " +
                "public void Write(BigExcelWriter writer) { const string sheet = \"Bad/Name\"; " +
                "writer.CreateAndOpenSheet(sheet); _ = new CellRange(\"A0\"); } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.Multiple(() =>
            {
                Assert.That(diagnostics.Count(item => item.Id == "BIGEXCEL4001"), Is.EqualTo(2));
                Assert.That(diagnostics.Any(item => item.GetMessage().Contains("sheet names cannot contain")), Is.True);
                Assert.That(diagnostics.Any(item => item.GetMessage().Contains("row numbers must be between")), Is.True);
            });
        }

        [Test]
        public async Task ValidAndDynamicRangesAreNotDiagnosed()
        {
            const string source = "using BigExcelCreator.Ranges; public class Model { " +
                "private string GetRange() => \"A0\"; public void Write() { " +
                "_ = new CellRange(\"'Sheet Name'!$A$1:B2\"); _ = new CellRange(GetRange()); } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL4001"), Is.False);
        }
    }
}
