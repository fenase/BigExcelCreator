// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.Ranges.InvalidConstantRange
{
    internal static class InvalidConstantRangeDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL4001",
            "Invalid constant sheet name or range",
            "Constant {0} '{1}' is invalid because {2}; the writer will reject it or Excel may not open the generated workbook.",
            "Usage",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);
    }
}
