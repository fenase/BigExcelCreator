// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using System.Collections.Generic;
using System.Linq;

namespace BigExcelCreator.Analyzers
{
    internal static class AnalyzerUtilities
    {
        internal static bool HasAttribute(ISymbol symbol, string metadataName)
        {
            if (symbol.GetAttributes().Any(attribute =>
                attribute.AttributeClass is INamedTypeSymbol attributeType &&
                GetMetadataName(attributeType) == metadataName))
            {
                return true;
            }

            for (IPropertySymbol? overriddenProperty = (symbol as IPropertySymbol)?.OverriddenProperty;
                 overriddenProperty is not null;
                 overriddenProperty = overriddenProperty.OverriddenProperty)
            {
                if (overriddenProperty.GetAttributes().Any(attribute =>
                    attribute.AttributeClass is INamedTypeSymbol attributeType &&
                    GetMetadataName(attributeType) == metadataName))
                {
                    return true;
                }
            }

            return false;
        }

        internal static AttributeData? FindAttribute(ISymbol symbol, string metadataName)
        {
            return symbol.GetAttributes().FirstOrDefault(attribute =>
                attribute.AttributeClass is INamedTypeSymbol attributeType &&
                GetMetadataName(attributeType) == metadataName);
        }

        internal static string GetMetadataName(INamedTypeSymbol type)
        {
            Stack<string> typeNames = new();
            for (INamedTypeSymbol? current = type; current is not null; current = current.ContainingType)
            {
                typeNames.Push(current.MetadataName);
            }

            string namespaceName = type.ContainingNamespace.IsGlobalNamespace
                ? string.Empty
                : type.ContainingNamespace.ToDisplayString() + ".";
            return namespaceName + string.Join("+", typeNames);
        }

        internal static IEnumerable<IPropertySymbol> GetExportedProperties(INamedTypeSymbol type)
        {
            HashSet<string> seenProperties = new(System.StringComparer.Ordinal);
            for (INamedTypeSymbol? current = type; current is not null; current = current.BaseType)
            {
                foreach (IPropertySymbol property in current.GetMembers().OfType<IPropertySymbol>())
                {
                    bool hasPublicAccessor = property.GetMethod?.DeclaredAccessibility == Accessibility.Public ||
                        property.SetMethod?.DeclaredAccessibility == Accessibility.Public;
                    if (!hasPublicAccessor || HasAttribute(property, KnownTypeNames.ExcelIgnoreAttribute))
                    {
                        continue;
                    }

                    string signature = property.Name + "(" + string.Join(",", property.Parameters.Select(parameter =>
                        parameter.Type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat))) + ")";
                    if (seenProperties.Add(signature))
                    {
                        yield return property;
                    }
                }
            }
        }

        internal static string GetTypeName(ITypeSymbol type)
        {
            return type.ToDisplayString(SymbolDisplayFormat.CSharpErrorMessageFormat);
        }
    }
}
