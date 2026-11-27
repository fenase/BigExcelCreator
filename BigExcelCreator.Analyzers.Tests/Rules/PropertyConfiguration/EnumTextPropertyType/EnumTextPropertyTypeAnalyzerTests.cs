// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.PropertyConfiguration.EnumTextPropertyType
{
    [TestFixture]
    internal sealed class EnumTextPropertyTypeAnalyzerTests
    {
        [Test]
        public async Task EnumTextOnNonEnumPropertyIsAnError()
        {
            const string source = "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public class Model { [ExcelColumnType(CellDataType.EnumText)] public string Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL0006"), Is.True);
        }

        [Test]
        public async Task EnumAndNullableEnumPropertiesAreSupported()
        {
            const string source = "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public enum Choice { A } public class Model { " +
                "[ExcelColumnType(CellDataType.EnumText)] public Choice Value { get; set; } " +
                "[ExcelColumnType(CellDataType.EnumText)] public Choice? Optional { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL0006"), Is.False);
        }
    }
}
