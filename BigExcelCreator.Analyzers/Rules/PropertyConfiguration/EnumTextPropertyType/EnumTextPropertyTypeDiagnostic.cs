// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.EnumTextPropertyType
{
    internal static class EnumTextPropertyTypeDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL0006",
            "Enum text requires an enum property",
            "Property '{0}' uses EnumText with CLR type '{1}'; export will fail because only enum values can be formatted as enum text.",
            "Usage",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);
    }
}
