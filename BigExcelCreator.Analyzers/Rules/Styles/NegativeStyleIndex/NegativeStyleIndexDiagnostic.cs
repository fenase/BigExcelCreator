// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.Styles.NegativeStyleIndex
{
    internal static class NegativeStyleIndexDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL3001",
            "Style index cannot be negative",
            "Style index {0} on '{1}' is negative; sheet generation will reject it instead of applying the requested style.",
            "Usage",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);
    }
}
