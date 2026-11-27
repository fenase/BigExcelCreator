// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.Styles.NegativeStyleIndex
{
    [TestFixture]
    internal sealed class NegativeStyleIndexAnalyzerTests
    {
        [Test]
        public async Task NegativeStyleIndexIsAnError()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelStyleFormat(-1)] public string Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL3001"), Is.True);
        }

        [Test]
        public async Task NonNegativeStyleIndexIsAccepted()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelStyleFormat(0)] public string Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL3001"), Is.False);
        }
    }
}
