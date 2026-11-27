// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Immutable;
using System.Linq;

namespace BigExcelCreator.Analyzers.Rules.Styles.NegativeStyleIndex
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class NegativeStyleIndexAnalyzer : DiagnosticAnalyzer
    {
        private static readonly string[] FormatAttributeTypes =
        [
            KnownTypeNames.ExcelStyleFormatAttribute,
            KnownTypeNames.ExcelHeaderStyleFormatAttribute,
            KnownTypeNames.ExcelConditionalFormatCellIsAttribute,
            KnownTypeNames.ExcelConditionalFormatFormulaAttribute,
            KnownTypeNames.ExcelConditionalFormatDuplicatedValuesAttribute,
        ];

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [NegativeStyleIndexDiagnostic.Descriptor];

        public override void Initialize(AnalysisContext context)
        {
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            context.EnableConcurrentExecution();
            context.RegisterSymbolAction(AnalyzeSymbol, SymbolKind.Property, SymbolKind.NamedType);
        }

        private static void AnalyzeSymbol(SymbolAnalysisContext context)
        {
            foreach (AttributeData attribute in context.Symbol.GetAttributes())
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                if (attribute.AttributeClass is not INamedTypeSymbol attributeType ||
                    !FormatAttributeTypes.Contains(AnalyzerUtilities.GetMetadataName(attributeType)) ||
                    attribute.ConstructorArguments.Length == 0 ||
                    attribute.ConstructorArguments[0].Type?.SpecialType != SpecialType.System_Int32 ||
                    attribute.ConstructorArguments[0].Value is not int format || format >= 0)
                {
                    continue;
                }

                string owner = context.Symbol is IPropertySymbol property ? property.Name : context.Symbol.Name;
                context.ReportDiagnostic(Diagnostic.Create(
                    NegativeStyleIndexDiagnostic.Descriptor,
                    attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation()
                        ?? context.Symbol.Locations.FirstOrDefault(),
                    format,
                    owner));
            }
        }
    }
}
