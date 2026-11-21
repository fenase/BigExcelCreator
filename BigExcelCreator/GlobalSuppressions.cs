// This file is used by Code Analysis to maintain SuppressMessage
// attributes that are applied to this project.
// Project-level suppressions either have no target or are given
// a specific target and scoped to a namespace, type, member, etc.

using System.Diagnostics.CodeAnalysis;

[assembly: SuppressMessage("Performance", "CA1851:Possible multiple enumerations of 'IEnumerable' collection", Justification = "<Pending>", Scope = "member", Target = "~M:BigExcelCreator.BigExcelWriter.CreateSheetFromObject``1(System.Collections.Generic.IEnumerable{``0},System.String,DocumentFormat.OpenXml.Spreadsheet.SheetStateValues,System.Boolean,System.Boolean,System.Collections.Generic.IList{DocumentFormat.OpenXml.Spreadsheet.Column})")]
[assembly: SuppressMessage("Style", "IDE0057:Use range operator", Justification = "This is a multi-target library. Some frameworks don't support the range operator.", Scope = "member", Target = "~M:BigExcelCreator.Extensions.SheetExtensions.CreateValidatorListSheetName(System.String,System.String)~System.String")]
