// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Analyzers.Rules.PropertyConfiguration.MultipleExcelValidators;
using BigExcelCreator.ClassAttributes;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Diagnostics;
using System;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BigExcelCreator.Analyzers.Tests
{
    [TestFixture]
    internal sealed class MultipleExcelValidatorsAnalyzerTests
    {
        [Test]
        public async Task NoValidatorProducesNoDiagnostics()
        {
            ImmutableArray<Diagnostic> diagnostics = await AnalyzeAsync("public class Model { public string Value { get; set; } }");

            Assert.That(diagnostics, Is.Empty);
        }

        [Test]
        public async Task OneValidatorProducesNoDiagnostics()
        {
            ImmutableArray<Diagnostic> diagnostics = await AnalyzeAsync(
                "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public class Model { [ExcelIntegerValidator(1, ValidatorOperator.Equal)] public int Value { get; set; } }");

            Assert.That(diagnostics, Is.Empty);
        }

        [Test]
        public async Task TwoValidatorsOfSameTypeProducesDiagnostic()
        {
            ImmutableArray<Diagnostic> diagnostics = await AnalyzeAsync(
                "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public class Model { [ExcelIntegerValidator(1, ValidatorOperator.Equal)] " +
                "[ExcelIntegerValidator(2, ValidatorOperator.Equal)] public int Value { get; set; } }");

            AssertDiagnostic(diagnostics, 1);
        }

        [Test]
        public async Task TwoDifferentValidatorsProducesDiagnostic()
        {
            ImmutableArray<Diagnostic> diagnostics = await AnalyzeAsync(
                "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public class Model { [ExcelIntegerValidator(1, ValidatorOperator.Equal)] " +
                "[ExcelDecimalValidator(2.0, ValidatorOperator.Equal)] public decimal Value { get; set; } }");

            AssertDiagnostic(diagnostics, 1);
        }

        [Test]
        public async Task ThreeValidatorsProducesDiagnosticForSecondAndThird()
        {
            ImmutableArray<Diagnostic> diagnostics = await AnalyzeAsync(
                "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public class Model { [ExcelIntegerValidator(1, ValidatorOperator.Equal)] " +
                "[ExcelDecimalValidator(2.0, ValidatorOperator.Equal)] " +
                "[ExcelListFormulaValidator(\"=A1\")] public decimal Value { get; set; } }");

            AssertDiagnostic(diagnostics, 2);
        }

        [Test]
        public async Task AliasedAndFullyQualifiedValidatorsAreRecognized()
        {
            const string source = "using ExcelInteger = BigExcelCreator.ClassAttributes.ExcelIntegerValidatorAttribute; " +
                "using BigExcelCreator.Enums; public class Model { " +
                "[ExcelInteger(1, ValidatorOperator.Equal)] " +
                "[global::BigExcelCreator.ClassAttributes.ExcelDecimalValidator(2.0, ValidatorOperator.Equal)] " +
                "public decimal Value { get; set; } }";
            ImmutableArray<Diagnostic> diagnostics = await AnalyzeAsync(source);

            AssertDiagnostic(diagnostics, 1);
        }

        [Test]
        public async Task UnrelatedAttributeWithValidatorLikeNameIsIgnored()
        {
            ImmutableArray<Diagnostic> diagnostics = await AnalyzeAsync(
                "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "namespace Other { public sealed class ExcelIntegerValidatorAttribute : System.Attribute { } } " +
                "public class Model { [ExcelIntegerValidator(1, ValidatorOperator.Equal)] " +
                "[Other.ExcelIntegerValidator] public int Value { get; set; } }");

            Assert.That(diagnostics, Is.Empty);
        }

        [Test]
        public async Task DerivedValidatorAttributeAndInheritedInterfaceAreRecognized()
        {
            const string source = "using System; " +
                "namespace BigExcelCreator.ClassAttributes.Interfaces { public interface IExcelValidatorAttribute { } } " +
                "[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)] " +
                "public class ValidatorAttribute : Attribute, BigExcelCreator.ClassAttributes.Interfaces.IExcelValidatorAttribute { } " +
                "public sealed class DerivedValidatorAttribute : ValidatorAttribute { } " +
                "public class Model { [Validator] [DerivedValidator] public string Value { get; set; } }";
            ImmutableArray<Diagnostic> diagnostics = await AnalyzeAsync(source, includeRuntime: false);

            AssertDiagnostic(diagnostics, 1);
        }

        [Test]
        public async Task DiagnosticHasExactIdMessageSeverityAndSecondAttributeLocation()
        {
            const string source = "using BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums;\n" +
                "public class Model\n{\n" +
                "    [ExcelIntegerValidator(1, ValidatorOperator.Equal)]\n" +
                "    [ExcelDecimalValidator(2.0, ValidatorOperator.Equal)]\n" +
                "    public decimal Value { get; set; }\n}";
            ImmutableArray<Diagnostic> diagnostics = await AnalyzeAsync(source);
            Diagnostic diagnostic = diagnostics.Single();
            int expectedStart = source.IndexOf("[ExcelDecimalValidator", StringComparison.Ordinal) + 1;
            int expectedLength = "ExcelDecimalValidator(2.0, ValidatorOperator.Equal)".Length;

            Assert.Multiple(() =>
            {
                Assert.That(diagnostic.Id, Is.EqualTo("BIGEXCEL0001"));
                Assert.That(diagnostic.GetMessage(), Is.EqualTo(
                    "Property 'Value' has multiple Excel validators; the writer can apply only one rule per generated Excel column, so the resulting workbook would omit validations."));
                Assert.That(diagnostic.Severity, Is.EqualTo(DiagnosticSeverity.Error));
                Assert.That(diagnostic.Location.SourceSpan.Start, Is.EqualTo(expectedStart));
                Assert.That(diagnostic.Location.SourceSpan.Length, Is.EqualTo(expectedLength));
            });
        }

        [Test]
        public async Task GeneratedCodeProducesNoDiagnostics()
        {
            ImmutableArray<Diagnostic> diagnostics = await AnalyzeAsync(
                "// <auto-generated/>\nusing BigExcelCreator.ClassAttributes; using BigExcelCreator.Enums; " +
                "public class GeneratedModel { [ExcelIntegerValidator(1, ValidatorOperator.Equal)] " +
                "[ExcelDecimalValidator(2.0, ValidatorOperator.Equal)] public decimal Value { get; set; } }",
                filePath: "GeneratedModel.g.cs");

            Assert.That(diagnostics, Is.Empty);
        }

        private static void AssertDiagnostic(ImmutableArray<Diagnostic> diagnostics, int expectedCount)
        {
            Assert.That(diagnostics.Length, Is.EqualTo(expectedCount));
            Assert.That(diagnostics.Select(diagnostic => diagnostic.Id), Is.All.EqualTo("BIGEXCEL0001"));
        }

        private static async Task<ImmutableArray<Diagnostic>> AnalyzeAsync(
            string source,
            string filePath = "TestModel.cs",
            bool includeRuntime = true)
        {
            CSharpParseOptions parseOptions = new(LanguageVersion.Latest);
            SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(source, parseOptions, filePath);
            CSharpCompilation compilation = CSharpCompilation.Create(
                "AnalyzerTest",
                new[] { syntaxTree },
                GetReferences(includeRuntime),
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
            ImmutableArray<DiagnosticAnalyzer> analyzers = ImmutableArray.Create<DiagnosticAnalyzer>(
                new MultipleExcelValidatorsAnalyzer());

            return await compilation.WithAnalyzers(analyzers).GetAnalyzerDiagnosticsAsync();
        }

        private static MetadataReference[] GetReferences(bool includeRuntime)
        {
            string trustedAssemblies = (string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!;
            var paths = trustedAssemblies.Split(Path.PathSeparator).ToList();
            if (includeRuntime)
            {
                paths.Add(typeof(ExcelListValidatorAttribute).Assembly.Location);
            }

            return paths.Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(path => MetadataReference.CreateFromFile(path))
                .ToArray();
        }
    }
}
