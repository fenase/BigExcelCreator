// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Operations;
using System;
using System.Collections.Immutable;
using System.Linq;
using System.Text.RegularExpressions;

namespace BigExcelCreator.Analyzers.Rules.Ranges.InvalidConstantRange
{
    [DiagnosticAnalyzer(LanguageNames.CSharp)]
    public sealed class InvalidConstantRangeAnalyzer : DiagnosticAnalyzer
    {
        private static readonly Regex RangeComponentPattern = new(
            @"^\$?(?:[A-Za-z]+(?:\$?\d+)?|\d+)$",
            RegexOptions.CultureInvariant,
            TimeSpan.FromSeconds(2));

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => [InvalidConstantRangeDiagnostic.Descriptor];

        public override void Initialize(AnalysisContext context)
        {
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
            context.EnableConcurrentExecution();
            context.RegisterOperationAction(AnalyzeInvocation, OperationKind.Invocation);
            context.RegisterOperationAction(AnalyzeObjectCreation, OperationKind.ObjectCreation);
        }

        private static void AnalyzeInvocation(OperationAnalysisContext context)
        {
            IInvocationOperation invocation = (IInvocationOperation)context.Operation;
            if (AnalyzerUtilities.GetMetadataName(invocation.TargetMethod.ContainingType) != KnownTypeNames.BigExcelWriter)
            {
                return;
            }

            foreach (IArgumentOperation argument in invocation.Arguments)
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                IParameterSymbol? parameter = argument.Parameter;
                if (parameter is null)
                {
                    continue;
                }

                bool isSheetName = invocation.TargetMethod.Name switch
                {
                    "CreateAndOpenSheet" => parameter.Name == "name",
                    "CreateSheetFromObject" => parameter.Name == "sheetName",
                    _ => false,
                };
                bool isRange = parameter.Name is "range" or "reference";
                if (!isSheetName && !isRange)
                {
                    continue;
                }

                Optional<object?> constantValue = argument.Value.ConstantValue;
                if (!constantValue.HasValue || constantValue.Value is not string value)
                {
                    continue;
                }

                string? reason = isSheetName ? GetInvalidSheetNameReason(value) : GetInvalidRangeReason(value);
                if (reason is not null)
                {
                    Report(context, argument.Syntax.GetLocation(), isSheetName ? "sheet name" : "range", value, reason);
                }
            }
        }

        private static void AnalyzeObjectCreation(OperationAnalysisContext context)
        {
            IObjectCreationOperation creation = (IObjectCreationOperation)context.Operation;
            if (creation.Constructor is null ||
                AnalyzerUtilities.GetMetadataName(creation.Constructor.ContainingType) != KnownTypeNames.CellRange)
            {
                return;
            }

            foreach (IArgumentOperation argument in creation.Arguments)
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                if (argument.Parameter?.Name != "range")
                {
                    continue;
                }

                Optional<object?> constantValue = argument.Value.ConstantValue;
                if (constantValue.HasValue && constantValue.Value is string value && GetInvalidRangeReason(value) is string reason)
                {
                    Report(context, argument.Syntax.GetLocation(), "range", value, reason);
                }
            }
        }

        private static string? GetInvalidSheetNameReason(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "Excel sheet names cannot be empty";
            }

            string trimmed = value.Trim();
            if (trimmed.Length > 31)
            {
                return "Excel sheet names cannot exceed 31 characters";
            }

            if (trimmed.Any(character => character < ' ' || "\\/*[]:?".IndexOf(character) >= 0))
            {
                return "Excel sheet names cannot contain control characters or \\ / * [ ] : ?";
            }

            return null;
        }

        private static string? GetInvalidRangeReason(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "cell references cannot be empty";
            }

            string range = value.Trim();
            int separator = range.IndexOf('!');
            if (separator >= 0)
            {
                if (separator != range.LastIndexOf('!'))
                {
                    return "a range can contain only one sheet separator";
                }

                string sheetName = range.Substring(0, separator);
                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    return "a sheet-qualified range must specify a sheet name";
                }

                if (GetInvalidSheetNameReason(sheetName) is string sheetReason)
                {
                    return sheetReason;
                }

                range = range.Substring(separator + 1);
            }

            string[] components = range.Split(':');
            if (components.Length > 2 || components.Any(component => !RangeComponentPattern.IsMatch(component)))
            {
                return "the range must contain valid Excel cell, row, or column references";
            }

            (string column, string row) first = GetColumnAndRow(components[0]);
            (string column, string row) last = GetColumnAndRow(components[components.Length - 1]);
            if ((first.column.Length == 0) != (last.column.Length == 0) ||
                (first.row.Length == 0) != (last.row.Length == 0))
            {
                return "both ends of a range must use the same cell, whole-row, or whole-column form";
            }

            if (first.column.Length > 0 &&
                (GetColumnIndex(first.column) > 16384 || GetColumnIndex(last.column) > 16384))
            {
                return "Excel supports columns only through XFD";
            }

            if (first.row.Length > 0 &&
                (!IsValidRow(first.row) || !IsValidRow(last.row)))
            {
                return "Excel row numbers must be between 1 and 1048576";
            }

            return null;
        }

        private static (string column, string row) GetColumnAndRow(string component)
        {
            string normalized = component.Replace("$", string.Empty);
            int firstDigit = normalized.IndexOfAny("0123456789".ToCharArray());
            return firstDigit < 0
                ? (normalized, string.Empty)
                : (normalized.Substring(0, firstDigit), normalized.Substring(firstDigit));
        }

        private static int GetColumnIndex(string column)
        {
            int index = 0;
            foreach (char character in column)
            {
                index = (index * 26) + (char.ToUpperInvariant(character) - 'A' + 1);
                if (index > 16384)
                {
                    return index;
                }
            }

            return index;
        }

        private static bool IsValidRow(string row)
        {
            return int.TryParse(row, out int rowNumber) && rowNumber is >= 1 and <= 1048576;
        }

        private static void Report(
            OperationAnalysisContext context,
            Location location,
            string kind,
            string value,
            string reason)
        {
            context.ReportDiagnostic(Diagnostic.Create(
                InvalidConstantRangeDiagnostic.Descriptor,
                location,
                kind,
                value,
                reason));
        }
    }
}
