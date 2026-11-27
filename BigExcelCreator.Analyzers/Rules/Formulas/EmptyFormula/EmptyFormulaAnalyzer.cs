// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Immutable;
using System.Linq;

namespace BigExcelCreator.Analyzers.Rules.Formulas.EmptyFormula
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class EmptyFormulaAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [EmptyFormulaDiagnostic.Descriptor];

        public override void Initialize(AnalysisContext context)
        {
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            context.EnableConcurrentExecution();
            context.RegisterSymbolAction(AnalyzeProperty, SymbolKind.Property);
        }

        private static void AnalyzeProperty(SymbolAnalysisContext context)
        {
            if (context.Symbol is not IPropertySymbol property ||
                AnalyzerUtilities.HasAttribute(property, KnownTypeNames.ExcelIgnoreAttribute))
            {
                return;
            }

            foreach (AttributeData attribute in property.GetAttributes())
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                if (attribute.AttributeClass is not INamedTypeSymbol attributeType)
                {
                    continue;
                }

                string metadataName = AnalyzerUtilities.GetMetadataName(attributeType);
                int formulaIndex = metadataName switch
                {
                    KnownTypeNames.ExcelListFormulaValidatorAttribute => 0,
                    KnownTypeNames.ExcelConditionalFormatFormulaAttribute => 1,
                    _ => -1,
                };
                if (formulaIndex < 0 || attribute.ConstructorArguments.Length <= formulaIndex)
                {
                    continue;
                }

                object? formulaValue = attribute.ConstructorArguments[formulaIndex].Value;
                if (formulaValue is string formula && !string.IsNullOrWhiteSpace(formula) ||
                    formulaValue is not null && formulaValue is not string)
                {
                    continue;
                }

                context.ReportDiagnostic(Diagnostic.Create(
                    EmptyFormulaDiagnostic.Descriptor,
                    attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation()
                        ?? property.Locations.FirstOrDefault(),
                    property.Name));
            }
        }
    }
}
