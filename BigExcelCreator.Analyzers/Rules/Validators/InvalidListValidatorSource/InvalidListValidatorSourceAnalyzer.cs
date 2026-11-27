// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using System.Collections.Immutable;
using System.Linq;

namespace BigExcelCreator.Analyzers.Rules.Validators.InvalidListValidatorSource
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class InvalidListValidatorSourceAnalyzer : DiagnosticAnalyzer
    {
        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [InvalidListValidatorSourceDiagnostic.Descriptor];

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

            AttributeData? attribute = AnalyzerUtilities.FindAttribute(property, KnownTypeNames.ExcelListValidatorAttribute);
            if (attribute?.AttributeConstructor is null)
            {
                return;
            }

            int typeIndex = GetParameterIndex(attribute, "enumType");
            if (typeIndex >= 0)
            {
                if (GetTypeArgument(attribute, typeIndex) is not { TypeKind: TypeKind.Enum })
                {
                    Report(context, property, attribute, "enumType");
                }

                return;
            }

            typeIndex = GetParameterIndex(attribute, "sourceType");
            int memberIndex = GetParameterIndex(attribute, "memberName");
            if (typeIndex < 0 || memberIndex < 0)
            {
                return;
            }

            if (GetTypeArgument(attribute, typeIndex) is not INamedTypeSymbol sourceType ||
                memberIndex >= attribute.ConstructorArguments.Length ||
                attribute.ConstructorArguments[memberIndex].Value is not string memberName ||
                string.IsNullOrWhiteSpace(memberName))
            {
                Report(context, property, attribute, "source member");
                return;
            }

            ISymbol? sourceMember = FindSourceMember(sourceType, memberName);
            if (sourceMember is null || !sourceMember.IsStatic ||
                GetMemberReturnType(sourceMember) is not ITypeSymbol returnType ||
                IsDefinitelyIncompatibleEnumerable(returnType))
            {
                Report(context, property, attribute, sourceType.Name + "." + memberName);
            }
        }

        private static int GetParameterIndex(AttributeData attribute, string parameterName)
        {
            for (int index = 0; index < attribute.AttributeConstructor!.Parameters.Length; index++)
            {
                if (attribute.AttributeConstructor.Parameters[index].Name == parameterName)
                {
                    return index;
                }
            }

            return -1;
        }

        private static INamedTypeSymbol? GetTypeArgument(AttributeData attribute, int index)
        {
            return index < attribute.ConstructorArguments.Length
                ? attribute.ConstructorArguments[index].Value as INamedTypeSymbol
                : null;
        }

        private static ISymbol? FindSourceMember(INamedTypeSymbol sourceType, string name)
        {
            for (INamedTypeSymbol? current = sourceType; current is not null; current = current.BaseType)
            {
                ISymbol? member = current.GetMembers(name).FirstOrDefault(candidate =>
                    candidate switch
                    {
                        IFieldSymbol field => field.IsStatic,
                        IPropertySymbol property => property.IsStatic && property.GetMethod is not null,
                        IMethodSymbol method => method.IsStatic && method.Parameters.Length == 0 && !method.IsGenericMethod,
                        _ => false,
                    });
                if (member is not null)
                {
                    return member;
                }
            }

            return null;
        }

        private static ITypeSymbol? GetMemberReturnType(ISymbol member)
        {
            return member switch
            {
                IFieldSymbol field => field.Type,
                IPropertySymbol property => property.Type,
                IMethodSymbol method => method.ReturnType,
                _ => null,
            };
        }

        private static bool IsDefinitelyIncompatibleEnumerable(ITypeSymbol type)
        {
            if (type.SpecialType == SpecialType.System_String || type.TypeKind == TypeKind.Error)
            {
                return false;
            }

            if (type is IArrayTypeSymbol arrayType)
            {
                return !IsStringOrEnum(arrayType.ElementType);
            }

            if (type is not INamedTypeSymbol namedType)
            {
                return false;
            }

            INamedTypeSymbol? enumerableOfT = namedType.AllInterfaces
                .Concat(new[] { namedType })
                .FirstOrDefault(candidate =>
                    candidate.OriginalDefinition.MetadataName == "IEnumerable`1" &&
                    candidate.OriginalDefinition.ContainingNamespace.ToDisplayString() == "System.Collections.Generic");
            if (enumerableOfT is not null)
            {
                return !IsStringOrEnum(enumerableOfT.TypeArguments[0]);
            }

            bool implementsNonGenericEnumerable = namedType.MetadataName == "IEnumerable" &&
                namedType.ContainingNamespace.ToDisplayString() == "System.Collections" ||
                namedType.AllInterfaces.Any(candidate =>
                    candidate.MetadataName == "IEnumerable" &&
                    candidate.ContainingNamespace.ToDisplayString() == "System.Collections");
            return !implementsNonGenericEnumerable && namedType.SpecialType != SpecialType.System_Object;
        }

        private static bool IsStringOrEnum(ITypeSymbol type)
        {
            if (type.SpecialType == SpecialType.System_String || type.TypeKind == TypeKind.Enum)
            {
                return true;
            }

            return type is INamedTypeSymbol nullableType &&
                nullableType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T &&
                nullableType.TypeArguments[0].TypeKind == TypeKind.Enum;
        }

        private static void Report(
            SymbolAnalysisContext context,
            IPropertySymbol property,
            AttributeData attribute,
            string sourceDescription)
        {
            context.ReportDiagnostic(Diagnostic.Create(
                InvalidListValidatorSourceDiagnostic.Descriptor,
                attribute.ApplicationSyntaxReference?.GetSyntax(context.CancellationToken).GetLocation()
                    ?? property.Locations.FirstOrDefault(),
                property.Name,
                sourceDescription));
        }
    }
}
