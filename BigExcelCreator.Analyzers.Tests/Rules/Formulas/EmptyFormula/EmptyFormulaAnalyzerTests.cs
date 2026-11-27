// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.Formulas.EmptyFormula
{
    [TestFixture]
    internal sealed class EmptyFormulaAnalyzerTests
    {
        [Test]
        public async Task EmptyListFormulaIsAnError()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelListFormulaValidator(\"  \" )] public string Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL2001"), Is.True);
        }

        [Test]
        public async Task NonEmptyFormulaIsAccepted()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelListFormulaValidator(\"NamedList\")] public string Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL2001"), Is.False);
        }
    }
}
