// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Analyzers.Rules.PropertyConfiguration.DuplicateColumnOrder;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Threading;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.DuplicateColumnOrder
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class DuplicateColumnOrderAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [DuplicateColumnOrderDiagnostic.Descriptor];

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

            Dictionary<int, IPropertySymbol> orders = [];
            foreach (IPropertySymbol property in AnalyzerUtilities.GetExportedProperties(type))
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                AttributeData? attribute = AnalyzerUtilities.FindAttribute(property, KnownTypeNames.ExcelColumnOrderAttribute);
                if (attribute is null || attribute.ConstructorArguments.Length == 0 ||
                    attribute.ConstructorArguments[0].Value is not int order)
                {
                    continue;
                }

                if (orders.TryGetValue(order, out IPropertySymbol? previousProperty))
                {
                    Location location = attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation()
                        ?? property.Locations.FirstOrDefault()
                        ?? Location.None;
                    context.ReportDiagnostic(Diagnostic.Create(
                        DuplicateColumnOrderDiagnostic.Descriptor,
                        location,
                        previousProperty.Name,
                        property.Name,
                        order));
                }
                else
                {
                    orders.Add(order, property);
                }
            }
        }
    }
}
