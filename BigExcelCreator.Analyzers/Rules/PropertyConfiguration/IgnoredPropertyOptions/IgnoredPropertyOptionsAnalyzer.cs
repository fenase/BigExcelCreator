// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Immutable;
using System.Linq;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.IgnoredPropertyOptions
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class IgnoredPropertyOptionsAnalyzer : DiagnosticAnalyzer
    {
        private static readonly string[] IgnoredPropertyOptionTypes =
        [
            KnownTypeNames.ExcelColumnNameAttribute,
            KnownTypeNames.ExcelColumnOrderAttribute,
            KnownTypeNames.ExcelColumnTypeAttribute,
            KnownTypeNames.ExcelColumnWidthAttribute,
            KnownTypeNames.ExcelColumnHiddenAttribute,
            KnownTypeNames.ExcelStyleFormatAttribute,
            KnownTypeNames.ExcelConditionalFormatCellIsAttribute,
            KnownTypeNames.ExcelConditionalFormatFormulaAttribute,
            KnownTypeNames.ExcelConditionalFormatDuplicatedValuesAttribute,
        ];

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [IgnoredPropertyOptionsDiagnostic.Descriptor];

        public override void Initialize(AnalysisContext context)
        {
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            context.EnableConcurrentExecution();
            context.RegisterCompilationStartAction(compilationContext =>
            {
                INamedTypeSymbol? validatorInterface = compilationContext.Compilation.GetTypeByMetadataName(
                    KnownTypeNames.ExcelValidatorAttributeInterface);
                compilationContext.RegisterSymbolAction(
                    symbolContext => AnalyzeProperty(symbolContext, validatorInterface),
                    SymbolKind.Property);
            });
        }

        private static void AnalyzeProperty(SymbolAnalysisContext context, INamedTypeSymbol? validatorInterface)
        {
            if (context.Symbol is not IPropertySymbol property ||
                !AnalyzerUtilities.HasAttribute(property, KnownTypeNames.ExcelIgnoreAttribute))
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
                bool isValidator = validatorInterface is not null && attributeType.ImplementsInterface(validatorInterface);
                if (!isValidator && !IgnoredPropertyOptionTypes.Contains(metadataName))
                {
                    continue;
                }

                context.ReportDiagnostic(Diagnostic.Create(
                    IgnoredPropertyOptionsDiagnostic.Descriptor,
                    attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation()
                        ?? property.Locations.FirstOrDefault(),
                    property.Name,
                    attributeType.Name));
            }
        }
    }
}
