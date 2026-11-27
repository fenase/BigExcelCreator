// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.Validators.MissingSecondOperand
{
    [TestFixture]
    internal sealed class MissingSecondOperandAnalyzerTests
    {
        [Test]
        public async Task BetweenValidatorWithoutSecondOperandWarnsAboutDefaultZero()
        {
            const string source = "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public class Model { [ExcelIntegerValidator(5, ValidatorOperator.Between)] public int Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);
            var diagnostic = diagnostics.Single(item => item.Id == "BIGEXCEL1001");

            Assert.Multiple(() =>
            {
                Assert.That(diagnostic.Severity, Is.EqualTo(Microsoft.CodeAnalysis.DiagnosticSeverity.Warning));
                Assert.That(diagnostic.GetMessage(), Does.Contain("defaults it to 0"));
            });
        }

        [Test]
        public async Task ExplicitSecondOperandDoesNotWarn()
        {
            const string source = "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public class Model { [ExcelIntegerValidator(5, ValidatorOperator.Between, secondOperand: 10)] " +
                "public int Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL1001"), Is.False);
        }
    }
}
