// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Analyzers.Rules.Formulas.EmptyFormula;
using BigExcelCreator.Analyzers.Rules.Formulas.InvalidConditionalFormatOperands;
using BigExcelCreator.Analyzers.Rules.PropertyConfiguration.DuplicateColumnName;
using BigExcelCreator.Analyzers.Rules.PropertyConfiguration.DuplicateColumnOrder;
using BigExcelCreator.Analyzers.Rules.PropertyConfiguration.EnumTextPropertyType;
using BigExcelCreator.Analyzers.Rules.PropertyConfiguration.IgnoredPropertyOptions;
using BigExcelCreator.Analyzers.Rules.PropertyConfiguration.MultipleExcelValidators;
using BigExcelCreator.Analyzers.Rules.PropertyConfiguration.UnexportableProperty;
using BigExcelCreator.Analyzers.Rules.Ranges.InvalidConstantRange;
using BigExcelCreator.Analyzers.Rules.Styles.NegativeStyleIndex;
using BigExcelCreator.Analyzers.Rules.Validators.InvalidListValidatorSource;
using BigExcelCreator.Analyzers.Rules.Validators.MissingSecondOperand;
using BigExcelCreator.ClassAttributes;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Diagnostics;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BigExcelCreator.Analyzers.Tests
{
    internal static class AnalyzerTestHelper
    {
        private static readonly DiagnosticAnalyzer[] Analyzers =
        [
            new MultipleExcelValidatorsAnalyzer(),
            new DuplicateColumnOrderAnalyzer(),
            new DuplicateColumnNameAnalyzer(),
            new UnexportablePropertyAnalyzer(),
            new IgnoredPropertyOptionsAnalyzer(),
            new EnumTextPropertyTypeAnalyzer(),
            new MissingSecondOperandAnalyzer(),
            new InvalidListValidatorSourceAnalyzer(),
            new EmptyFormulaAnalyzer(),
            new InvalidConditionalFormatOperandsAnalyzer(),
            new NegativeStyleIndexAnalyzer(),
            new InvalidConstantRangeAnalyzer(),
        ];

        internal static async Task<ImmutableArray<Diagnostic>> AnalyzeAsync(string source, string filePath = "TestModel.cs")
        {
            CSharpParseOptions parseOptions = new(LanguageVersion.Latest);
            SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(source, parseOptions, filePath);
            CSharpCompilation compilation = CSharpCompilation.Create(
                "AnalyzerTest",
                new[] { syntaxTree },
                GetReferences(),
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

            return await compilation.WithAnalyzers(ImmutableArray.Create(Analyzers)).GetAnalyzerDiagnosticsAsync();
        }

        private static IEnumerable<MetadataReference> GetReferences()
        {
            string trustedAssemblies = (string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!;
            IEnumerable<string> paths = trustedAssemblies.Split(Path.PathSeparator)
                .Concat(new[]
                {
                    typeof(ExcelListValidatorAttribute).Assembly.Location,
                    typeof(SpreadsheetDocument).Assembly.Location,
                })
                .Distinct(StringComparer.OrdinalIgnoreCase);

            return paths.Select(path => MetadataReference.CreateFromFile(path));
        }
    }
}
