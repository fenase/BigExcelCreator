// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.Formulas.EmptyFormula
{
    internal static class EmptyFormulaDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL2001",
            "Formula must not be empty",
            "Formula on property '{0}' is empty; Excel cannot apply the requested validation or formatting, and sheet generation may fail.",
            "Usage",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);
    }
}
