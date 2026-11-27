// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.Validators.MissingSecondOperand
{
    internal static class MissingSecondOperandDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL1001",
            "Missing second validator operand",
            "Validator on property '{0}' uses '{1}' without secondOperand; the attribute defaults it to 0, which can produce an unintended Excel validation range.",
            "Usage",
            DiagnosticSeverity.Warning,
            isEnabledByDefault: true);
    }
}
