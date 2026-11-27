// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.DuplicateColumnName
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class DuplicateColumnNameAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [DuplicateColumnNameDiagnostic.Descriptor];

        public override void Initialize(AnalysisContext context)
        {
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            context.EnableConcurrentExecution();
            context.RegisterSymbolAction(AnalyzeType, SymbolKind.NamedType);
        }

        private static void AnalyzeType(SymbolAnalysisContext context)
        {
            if (context.Symbol is not INamedTypeSymbol type || type.TypeKind != TypeKind.Class)
            {
                return;
            }

            Dictionary<string, IPropertySymbol> names = new(System.StringComparer.OrdinalIgnoreCase);
            foreach (IPropertySymbol property in AnalyzerUtilities.GetExportedProperties(type))
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                AttributeData? attribute = AnalyzerUtilities.FindAttribute(property, KnownTypeNames.ExcelColumnNameAttribute);
                string columnName = attribute?.ConstructorArguments.FirstOrDefault().Value as string ?? property.Name;

                if (names.TryGetValue(columnName, out IPropertySymbol? previousProperty))
                {
                    Location location = attribute?.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation()
                        ?? property.Locations.FirstOrDefault()
                        ?? Location.None;
                    context.ReportDiagnostic(Diagnostic.Create(
                        DuplicateColumnNameDiagnostic.Descriptor,
                        location,
                        previousProperty.Name,
                        property.Name,
                        columnName));
                }
                else
                {
                    names.Add(columnName, property);
                }
            }
        }
    }
}
