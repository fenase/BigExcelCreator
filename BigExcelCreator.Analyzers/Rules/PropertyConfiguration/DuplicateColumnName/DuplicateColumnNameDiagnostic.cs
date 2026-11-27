// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.DuplicateColumnName
{
    internal static class DuplicateColumnNameDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL0003",
            "Duplicate Excel column name",
            "Properties '{0}' and '{1}' produce the same Excel header '{2}'; users may not be able to distinguish the columns in the generated sheet.",
            "Usage",
            DiagnosticSeverity.Warning,
            isEnabledByDefault: true);
    }
}
