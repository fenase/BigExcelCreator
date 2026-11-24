// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

#if NET48_OR_GREATER || NET6_0_OR_GREATER || NET8_0_OR_GREATER || NET10_0_OR_GREATER
using BigExcelCreator.ClassAttributes;
using BigExcelCreator.Enums;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace BigExcelCreator.Test
{
    [TestFixture]
    internal sealed class ExcelListValidatorAttributeDisplayTests
    {
        internal static readonly string[] expectedWhenUseDisplayName = ["Yes, eligible", "No, not eligible", "Pending"];
        internal static readonly string[] expectedWhenUseDisplayNameIsFalse = ["Yes", "No", "Pending"];

        [Test]
        public void EnumConstructorResolvesDisplayNamesAndFallbacks()
        {
            ExcelListValidatorAttribute attribute = new(typeof(DisplayTestEnumExample));
            ExcelListValidatorAttribute rawAttribute = new(typeof(DisplayTestEnumExample), false);

            Assert.Multiple(() =>
            {
                Assert.That(attribute.EnumType, Is.EqualTo(typeof(DisplayTestEnumExample)));
                Assert.That(attribute.UseDisplayName, Is.True);
                Assert.That(attribute.Values, Is.EqualTo(expectedWhenUseDisplayName));
                Assert.That(rawAttribute.EnumType, Is.EqualTo(typeof(DisplayTestEnumExample)));
                Assert.That(rawAttribute.UseDisplayName, Is.False);
                Assert.That(rawAttribute.Values, Is.EqualTo(expectedWhenUseDisplayNameIsFalse));
            });
        }

        [Test]
        public void EnumConstructorWithDisplayOptionsUsesResolvedNamesAndFlags()
        {
            ExcelListValidatorAttribute attribute = new(false, false, false, typeof(DisplayTestEnumExample));
            ExcelListValidatorAttribute rawAttribute = new(false, false, false, typeof(DisplayTestEnumExample), false);

            Assert.Multiple(() =>
            {
                Assert.That(attribute.EnumType, Is.EqualTo(typeof(DisplayTestEnumExample)));
                Assert.That(attribute.UseDisplayName, Is.True);
                Assert.That(attribute.Values, Is.EqualTo(expectedWhenUseDisplayName));
                Assert.That(rawAttribute.EnumType, Is.EqualTo(typeof(DisplayTestEnumExample)));
                Assert.That(rawAttribute.UseDisplayName, Is.False);
                Assert.That(rawAttribute.Values, Is.EqualTo(expectedWhenUseDisplayNameIsFalse));
                Assert.That(attribute.AllowBlank, Is.False);
                Assert.That(attribute.ShowInputMessage, Is.False);
                Assert.That(attribute.ShowErrorMessage, Is.False);
            });
        }

        [Test]
        public void CreateSheetFromObjectWritesEnumTextDisplayNames()
        {
            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream))
            {
                writer.CreateSheetFromObject([new EnumTextModel()], "EnumTextDisplay");
            }

            using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = document.WorkbookPart!;
            Sheet targetSheet = workbookPart.Workbook!.Sheets!.ChildElements
                .OfType<Sheet>()
                .Single(sheet => sheet.Name?.Value == "EnumTextDisplay");
            WorksheetPart targetPart = (WorksheetPart)workbookPart.GetPartById(targetSheet.Id!.Value!);
            Row[] rows = [.. targetPart.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>()];
            string[] values = [.. rows[1].Elements<Cell>().Select(cell => TestHelperMethods.GetCellRealValue(cell, workbookPart))];

            Assert.That(values, Is.EqualTo(expectedWhenUseDisplayName));
        }

        [Test]
        public void MemberConstructorResolvesEnumDisplayNames()
        {
            ExcelListValidatorAttribute attribute = new(typeof(DisplaySource), nameof(DisplaySource.Values));
            ExcelListValidatorAttribute rawAttribute = new(typeof(DisplaySource), nameof(DisplaySource.Values), false);

            Assert.Multiple(() =>
            {
                Assert.That(attribute.SourceType, Is.EqualTo(typeof(DisplaySource)));
                Assert.That(attribute.MemberName, Is.EqualTo(nameof(DisplaySource.Values)));
                Assert.That(attribute.UseDisplayName, Is.True);
                Assert.That(attribute.Values, Is.EqualTo(expectedWhenUseDisplayName));
                Assert.That(rawAttribute.UseDisplayName, Is.False);
                Assert.That(rawAttribute.Values, Is.EqualTo(expectedWhenUseDisplayNameIsFalse));
            });
        }

        private enum DisplayTestEnumExample
        {
            [Display(Name = "Yes, eligible")]
            Yes,

            [Display(Name = "No, not eligible")]
            No,

            Pending
        }

        private sealed class EnumTextModel
        {
            [ExcelColumnType(CellDataType.EnumText)]
            public DisplayTestEnumExample Display { get; set; } = DisplayTestEnumExample.Yes;

            [ExcelColumnType(CellDataType.EnumText)]
            public DisplayTestEnumExample NotEligible { get; set; } = DisplayTestEnumExample.No;

            [ExcelColumnType(CellDataType.EnumText)]
            public DisplayTestEnumExample PendingValue { get; set; } = DisplayTestEnumExample.Pending;
        }

        private static class DisplaySource
        {
            public static IEnumerable<Enum> Values
            {
                get
                {
                    return [DisplayTestEnumExample.Yes, DisplayTestEnumExample.No, DisplayTestEnumExample.Pending];
                }
            }
        }
    }
}
#endif
