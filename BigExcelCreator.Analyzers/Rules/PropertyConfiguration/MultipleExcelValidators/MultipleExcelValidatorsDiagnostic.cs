// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.MultipleExcelValidators
{
    internal static class MultipleExcelValidatorsDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL0001",
            "Multiple Excel validators",
            "Property '{0}' has multiple Excel validators; the writer can apply only one rule per generated Excel column, so the resulting workbook would omit validations.",
            "Usage",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);
    }
}
