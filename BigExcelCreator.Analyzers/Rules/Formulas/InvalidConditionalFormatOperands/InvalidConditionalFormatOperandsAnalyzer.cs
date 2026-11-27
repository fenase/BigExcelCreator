// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Immutable;
using System.Linq;

namespace BigExcelCreator.Analyzers.Rules.Formulas.InvalidConditionalFormatOperands
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class InvalidConditionalFormatOperandsAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [InvalidConditionalFormatOperandsDiagnostic.Descriptor];

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

            AttributeData? attribute = AnalyzerUtilities.FindAttribute(
                property,
                KnownTypeNames.ExcelConditionalFormatCellIsAttribute);
            if (attribute is null || attribute.ConstructorArguments.Length < 4 ||
                attribute.ConstructorArguments[1].Type is not INamedTypeSymbol operatorType)
            {
                return;
            }

            IFieldSymbol? operatorField = operatorType.GetMembers().OfType<IFieldSymbol>()
                .FirstOrDefault(field => field.HasConstantValue && Equals(field.ConstantValue, attribute.ConstructorArguments[1].Value));
            if (operatorField is null)
            {
                return;
            }

            string? invalidOperand = attribute.ConstructorArguments[2].Value is not string firstValue ||
                string.IsNullOrWhiteSpace(firstValue)
                ? "the first comparison value is empty"
                : null;

            if (invalidOperand is null &&
                (operatorField.Name == "Between" || operatorField.Name == "NotBetween") &&
                (attribute.ConstructorArguments[3].Value is not string secondValue || string.IsNullOrWhiteSpace(secondValue)))
            {
                invalidOperand = "a second comparison value is required but empty";
            }

            if (invalidOperand is null)
            {
                return;
            }

            context.ReportDiagnostic(Diagnostic.Create(
                InvalidConditionalFormatOperandsDiagnostic.Descriptor,
                attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation()
                    ?? property.Locations.FirstOrDefault(),
                property.Name,
                operatorField.Name,
                invalidOperand));
        }
    }
}
