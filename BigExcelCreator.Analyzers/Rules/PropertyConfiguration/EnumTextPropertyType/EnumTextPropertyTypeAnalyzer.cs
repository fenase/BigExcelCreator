// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Immutable;
using System.Linq;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.EnumTextPropertyType
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class EnumTextPropertyTypeAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [EnumTextPropertyTypeDiagnostic.Descriptor];

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

            AttributeData? attribute = AnalyzerUtilities.FindAttribute(property, KnownTypeNames.ExcelColumnTypeAttribute);
            if (attribute is null || attribute.ConstructorArguments.Length == 0)
            {
                return;
            }

            TypedConstant cellType = attribute.ConstructorArguments[0];
            if (cellType.Type is not INamedTypeSymbol enumType || enumType.TypeKind != TypeKind.Enum)
            {
                return;
            }

            IFieldSymbol? enumText = enumType.GetMembers("EnumText").OfType<IFieldSymbol>().FirstOrDefault();
            if (enumText is null || !Equals(enumText.ConstantValue, cellType.Value))
            {
                return;
            }

            ITypeSymbol propertyType = property.Type;
            if (propertyType is INamedTypeSymbol nullableType &&
                nullableType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
            {
                propertyType = nullableType.TypeArguments[0];
            }

            if (propertyType.TypeKind != TypeKind.Enum)
            {
                context.ReportDiagnostic(Diagnostic.Create(
                    EnumTextPropertyTypeDiagnostic.Descriptor,
                    attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation()
                        ?? property.Locations.FirstOrDefault(),
                    property.Name,
                    AnalyzerUtilities.GetTypeName(property.Type)));
            }
        }
    }
}
