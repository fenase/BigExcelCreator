// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers
{
    internal static class DiagnosticDescriptors
    {
        internal static readonly DiagnosticDescriptor MultipleExcelValidators = new(
            DiagnosticIds.MultipleExcelValidators,
            "Multiple Excel validators",
            "Property '{0}' has multiple Excel validators; only one validator is allowed per generated Excel column.",
            "Usage",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);
    }
}
