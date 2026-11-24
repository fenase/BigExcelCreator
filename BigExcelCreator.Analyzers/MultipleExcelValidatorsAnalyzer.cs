// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Threading;

namespace BigExcelCreator.Analyzers
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class MultipleExcelValidatorsAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
            ImmutableArray.Create(DiagnosticDescriptors.MultipleExcelValidators);

        public override void Initialize(AnalysisContext context)
        {
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            context.EnableConcurrentExecution();
            context.RegisterSymbolAction(AnalyzeProperty, SymbolKind.Property);
        }

        private static void AnalyzeProperty(SymbolAnalysisContext context)
        {
            CancellationToken cancellationToken = context.CancellationToken;
            if (context.Symbol is not IPropertySymbol property)
            {
                return;
            }

            INamedTypeSymbol? validatorInterface = context.Compilation.GetTypeByMetadataName(
                KnownTypeNames.ExcelValidatorAttributeInterface);
            if (validatorInterface is null)
            {
                return;
            }

            List<AttributeData> validators = [];
            foreach (AttributeData attribute in property.GetAttributes())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (attribute.AttributeClass is INamedTypeSymbol attributeType &&
                    attributeType.ImplementsInterface(validatorInterface))
                {
                    validators.Add(attribute);
                }
            }

            for (int index = 1; index < validators.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Location location = validators[index].ApplicationSyntaxReference?
                    .GetSyntax(cancellationToken)
                    .GetLocation() ?? property.Locations.FirstOrDefault() ?? Location.None;
                context.ReportDiagnostic(Diagnostic.Create(
                    DiagnosticDescriptors.MultipleExcelValidators,
                    location,
                    property.Name));
            }
        }
    }
}
