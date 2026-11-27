// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.Validators.InvalidListValidatorSource
{
    internal static class InvalidListValidatorSourceDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL1002",
            "Invalid list validator source",
            "List validator on property '{0}' references '{1}', which is not an enum or a static member returning strings or enum values; sheet creation will fail at runtime.",
            "Usage",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);
    }
}
