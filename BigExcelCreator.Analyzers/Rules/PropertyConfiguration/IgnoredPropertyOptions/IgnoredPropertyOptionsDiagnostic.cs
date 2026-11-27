// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;

namespace BigExcelCreator.Analyzers.Rules.PropertyConfiguration.IgnoredPropertyOptions
{
    internal static class IgnoredPropertyOptionsDiagnostic
    {
        internal static readonly DiagnosticDescriptor Descriptor = new(
            "BIGEXCEL0005",
            "Configuration on ignored property",
            "Property '{0}' is ignored, so '{1}' will not be applied to the generated Excel sheet.",
            "Usage",
            DiagnosticSeverity.Info,
            isEnabledByDefault: true);
    }
}
