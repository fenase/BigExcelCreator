// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.PropertyConfiguration.DuplicateColumnOrder
{
    [TestFixture]
    internal sealed class DuplicateColumnOrderAnalyzerTests
    {
        [Test]
        public async Task DuplicateOrdersWarnThatTiedColumnsAreOrderedByPropertyName()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelColumnOrder(1)] public int First { get; set; } " +
                "[ExcelColumnOrder(1)] public int Second { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);
            var diagnostic = diagnostics.Single(item => item.Id == "BIGEXCEL0002");

            Assert.Multiple(() =>
            {
                Assert.That(diagnostic.Severity, Is.EqualTo(Microsoft.CodeAnalysis.DiagnosticSeverity.Warning));
                Assert.That(diagnostic.GetMessage(), Is.EqualTo(
                    "Properties 'First' and 'Second' share Excel column order 1; tied columns are ordered alphabetically by property name, which may differ from the intended column layout."));
            });
        }

        [Test]
        public async Task DifferentOrdersProduceNoDuplicateOrderDiagnostic()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelColumnOrder(1)] public int First { get; set; } " +
                "[ExcelColumnOrder(2)] public int Second { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL0002"), Is.False);
        }
    }
}
