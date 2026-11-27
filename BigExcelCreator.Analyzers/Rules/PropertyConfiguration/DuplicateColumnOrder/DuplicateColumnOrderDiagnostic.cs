// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.DuplicateColumnOrder
{
    internal static class DuplicateColumnOrderDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL0002",
            "Duplicate Excel column order",
            "Properties '{0}' and '{1}' share Excel column order {2}; tied columns are ordered alphabetically by property name, which may differ from the intended column layout.",
            "Usage",
            DiagnosticSeverity.Warning,
            isEnabledByDefault: true);
    }
}
