// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests.Rules.PropertyConfiguration.UnexportableProperty
{
    [TestFixture]
    internal sealed class UnexportablePropertyAnalyzerTests
    {
        [Test]
        public async Task WriteOnlyAndIndexerPropertiesAreErrors()
        {
            const string source = "public class Model { public string WriteOnly { set { } } " +
                "public string this[int index] { get { return string.Empty; } } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Count(item => item.Id == "BIGEXCEL0004"), Is.EqualTo(2));
        }

        [Test]
        public async Task ReadablePropertyProducesNoDiagnostic()
        {
            const string source = "public class Model { public string Value { get; set; } }";
            var diagnostics = await AnalyzerTestHelper.AnalyzeAsync(source);

            Assert.That(diagnostics.Any(item => item.Id == "BIGEXCEL0004"), Is.False);
        }
    }
}
