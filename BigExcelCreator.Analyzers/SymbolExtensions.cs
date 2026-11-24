// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using System.Linq;

namespace BigExcelCreator.Analyzers
{
    internal static class SymbolExtensions
    {
        internal static bool ImplementsInterface(this INamedTypeSymbol type, INamedTypeSymbol interfaceType)
        {
            return type.AllInterfaces.Any(implementedInterface =>
                SymbolEqualityComparer.Default.Equals(implementedInterface, interfaceType));
        }
    }
}
