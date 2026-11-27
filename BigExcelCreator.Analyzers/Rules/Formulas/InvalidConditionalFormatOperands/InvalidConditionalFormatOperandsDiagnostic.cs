// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.Formulas.InvalidConditionalFormatOperands
{
    internal static class InvalidConditionalFormatOperandsDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL2002",
            "Conditional-format operands are invalid",
            "Conditional formatting on property '{0}' uses operator '{1}' but {2}; the writer will reject the rule, so Excel will not apply the requested formatting.",
            "Usage",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);
    }
}
