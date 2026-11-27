// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.Validators.InvalidListValidatorSource
{
    [TestFixture]
    internal sealed class InvalidListValidatorSourceAnalyzerTests
    {
        [Test]
        public async Task NonEnumTypeForEnumSourceIsAnError()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelListValidator(typeof(string))] public string Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL1002"), Is.True);
        }

        [Test]
        public async Task MissingStaticSourceMemberIsAnError()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Source { } public class Model { " +
                "[ExcelListValidator(typeof(Source), \"Missing\")] public string Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL1002"), Is.True);
        }

        [Test]
        public async Task ValidEnumAndStringSourceAreAccepted()
        {
            const string source = "using System.Collections.Generic; using BigExcelCreator.ClassAttributes; " +
                "public enum Choice { A } public class Source { public static IEnumerable<string> Values => null; } " +
                "public class Model { [ExcelListValidator(typeof(Choice))] public Choice EnumValue { get; set; } " +
                "[ExcelListValidator(typeof(Source), nameof(Source.Values))] public string TextValue { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL1002"), Is.False);
        }

        [Test]
        public async Task LiteralValuesAreNotTreatedAsStaticMemberSources()
        {
            const string source = "using BigExcelCreator.ClassAttributes; public class Model { " +
                "[ExcelListValidator(\"A\", \"B\")] public string PositionalValues { get; set; } " +
                "[ExcelListValidator(values: [\"C\", \"D\"])] public string NamedValues { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL1002"), Is.False);
        }
    }
}
