// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.PropertyConfiguration.DuplicateColumnName
{
    [TestFixture]
    internal sealed class DuplicateColumnNameAnalyzerTests
    {
        [Test]
        public async Task DuplicateHeadersWarnThatUsersCannotDistinguishColumns()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelColumnName(\"Status\")] public int First { get; set; } " +
                "[ExcelColumnName(\"status\")] public int Second { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL0003" &&
                item.GetMessage().Contains("users may not be able to distinguish the columns")), Is.True);
        }

        [Test]
        public async Task DistinctHeadersProduceNoDiagnostic()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelColumnName(\"First\")] public int First { get; set; } " +
                "[ExcelColumnName(\"Second\")] public int Second { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL0003"), Is.False);
        }
    }
}
