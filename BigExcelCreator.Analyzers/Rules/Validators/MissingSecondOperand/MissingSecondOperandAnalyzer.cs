// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Immutable;
using System.Linq;

namespace BigExcelCreator.Analyzers.Rules.Validators.MissingSecondOperand
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class MissingSecondOperandAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [MissingSecondOperandDiagnostic.Descriptor];

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
                if (attribute.AttributeClass is not INamedTypeSymbol attributeType ||
                    (AnalyzerUtilities.GetMetadataName(attributeType) != KnownTypeNames.ExcelIntegerValidatorAttribute &&
                     AnalyzerUtilities.GetMetadataName(attributeType) != KnownTypeNames.ExcelDecimalValidatorAttribute) ||
                    attribute.ConstructorArguments.Length < 2 ||
                    attribute.ConstructorArguments[1].Type is not INamedTypeSymbol operatorType)
                {
                    continue;
                }

                IFieldSymbol? operatorField = operatorType.GetMembers().OfType<IFieldSymbol>()
                    .FirstOrDefault(field => field.HasConstantValue && Equals(field.ConstantValue, attribute.ConstructorArguments[1].Value));
                if (operatorField?.Name != "Between" && operatorField?.Name != "NotBetween")
                {
                    continue;
                }

                AttributeSyntax? syntax = attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken) as AttributeSyntax;
                if (syntax?.ArgumentList is null)
                {
                    continue;
                }

                bool hasSecondOperand = syntax.ArgumentList.Arguments.Any(argument =>
                    argument.NameColon?.Name.Identifier.ValueText == "secondOperand") ||
                    syntax.ArgumentList.Arguments.Count(argument => argument.NameColon is null && argument.NameEquals is null) >= 6;
                if (hasSecondOperand)
                {
                    continue;
                }

                context.ReportDiagnostic(Diagnostic.Create(
                    MissingSecondOperandDiagnostic.Descriptor,
                    syntax.GetLocation(),
                    property.Name,
                    operatorField.Name));
            }
        }
    }
}
