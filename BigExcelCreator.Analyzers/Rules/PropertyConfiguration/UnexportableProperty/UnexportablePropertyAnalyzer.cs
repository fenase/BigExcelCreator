// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Immutable;
using System.Linq;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.UnexportableProperty
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class UnexportablePropertyAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [UnexportablePropertyDiagnostic.Descriptor];

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

            bool hasPublicAccessor = property.GetMethod?.DeclaredAccessibility == Accessibility.Public ||
                property.SetMethod?.DeclaredAccessibility == Accessibility.Public;
            if (!hasPublicAccessor)
            {
                return;
            }

            string? reason = property.Parameters.Length > 0
                ? "it is an indexer and the writer does not supply index arguments"
                : property.GetMethod is null
                    ? "it has no getter"
                    : null;
            if (reason is not null)
            {
                context.ReportDiagnostic(Diagnostic.Create(
                    UnexportablePropertyDiagnostic.Descriptor,
                    property.Locations.FirstOrDefault(),
                    property.Name,
                    reason));
            }
        }
    }
}
