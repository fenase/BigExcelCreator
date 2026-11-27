// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Threading;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.MultipleExcelValidators
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class MultipleExcelValidatorsAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
            [MultipleExcelValidatorsDiagnostic.Descriptor];

        public override void Initialize(AnalysisContext context)
        {
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            context.EnableConcurrentExecution();

            context.RegisterCompilationStartAction(compilationContext =>
            {
                INamedTypeSymbol? validatorInterface =
                    compilationContext.Compilation.GetTypeByMetadataName(
                        KnownTypeNames.ExcelValidatorAttributeInterface);

                if (validatorInterface is null)
                {
                    return;
                }

                compilationContext.RegisterSymbolAction(
                    symbolContext => AnalyzeProperty(
                        symbolContext,
                        validatorInterface),
                    SymbolKind.Property);
            });
        }

        private static void AnalyzeProperty(
            SymbolAnalysisContext context,
            INamedTypeSymbol validatorInterface)
        {
            CancellationToken cancellationToken = context.CancellationToken;

            if (context.Symbol is not IPropertySymbol property)
            {
                return;
            }

            List<AttributeData> directValidators = GetValidators(
                property,
                validatorInterface,
                cancellationToken);

            int inheritedValidatorCount = 0;

            for (IPropertySymbol? baseProperty = property.OverriddenProperty;
                 baseProperty is not null;
                 baseProperty = baseProperty.OverriddenProperty)
            {
                inheritedValidatorCount += GetValidators(
                    baseProperty,
                    validatorInterface,
                    cancellationToken).Count;
            }

            if (inheritedValidatorCount == 0)
            {
                for (int index = 1; index < directValidators.Count; index++)
                {
                    ReportDiagnostic(
                        context,
                        property,
                        directValidators[index],
                        cancellationToken);
                }

                return;
            }

            foreach (AttributeData directValidator in directValidators)
            {
                ReportDiagnostic(
                    context,
                    property,
                    directValidator,
                    cancellationToken);
            }
        }

        private static List<AttributeData> GetValidators(
            IPropertySymbol property,
            INamedTypeSymbol validatorInterface,
            CancellationToken cancellationToken)
        {
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

            return validators;
        }

        private static void ReportDiagnostic(
            SymbolAnalysisContext context,
            IPropertySymbol property,
            AttributeData validator,
            CancellationToken cancellationToken)
        {
            Location location =
                validator.ApplicationSyntaxReference?
                    .GetSyntax(cancellationToken)
                    .GetLocation()
                ?? property.Locations.FirstOrDefault()
                ?? Location.None;

            context.ReportDiagnostic(
                Diagnostic.Create(
                    MultipleExcelValidatorsDiagnostic.Descriptor,
                    location,
                    property.Name));
        }
    }
}