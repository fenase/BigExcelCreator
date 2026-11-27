// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.PropertyConfiguration.IgnoredPropertyOptions
{
    [TestFixture]
    internal sealed class IgnoredPropertyOptionsAnalyzerTests
    {
        [Test]
        public async Task IgnoredPropertyConfigurationIsInformational()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelIgnore] [ExcelColumnName(\"Hidden\")] public string Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);
            var diagnostic = diagnostics.Single(item => item.Id == "BIGEXCEL0005");

            Assert.That(diagnostic.Severity, Is.EqualTo(Microsoft.CodeAnalysis.DiagnosticSeverity.Info));
        }

        [Test]
        public async Task IgnoredPropertyWithoutOptionsProducesNoDiagnostic()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelIgnore] public string Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL0005"), Is.False);
        }
    }
}
