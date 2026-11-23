// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

#if NET48_OR_GREATER || NET6_0_OR_GREATER || NET8_0_OR_GREATER || NET10_0_OR_GREATER
using BigExcelCreator.ClassAttributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using BigExcelCreator.Enums;

namespace BigExcelCreator.Test
{
    [TestFixture]
    internal sealed class ExcelListValidatorAttributeDisplayTests
    {
        [Test]
        public void EnumConstructorResolvesDisplayNamesAndFallbacks()
        {
            ExcelListValidatorAttribute attribute = new ExcelListValidatorAttribute(typeof(DisplayTestEnum));
            ExcelListValidatorAttribute rawAttribute = new ExcelListValidatorAttribute(typeof(DisplayTestEnum), false);

            Assert.Multiple(() =>
            {
                Assert.That(attribute.EnumType, Is.EqualTo(typeof(DisplayTestEnum)));
                Assert.That(attribute.UseDisplayName, Is.True);
                Assert.That(attribute.Values, Is.EqualTo(new[] { "Yes, eligible", "No, not eligible", "Pending" }));
                Assert.That(rawAttribute.EnumType, Is.EqualTo(typeof(DisplayTestEnum)));
                Assert.That(rawAttribute.UseDisplayName, Is.False);
                Assert.That(rawAttribute.Values, Is.EqualTo(new[] { "Yes", "No", "Pending" }));
            });
        }

        [Test]
        public void EnumConstructorWithDisplayOptionsUsesResolvedNamesAndFlags()
        {
            ExcelListValidatorAttribute attribute = new ExcelListValidatorAttribute(false, false, false, typeof(DisplayTestEnum));
            ExcelListValidatorAttribute rawAttribute = new ExcelListValidatorAttribute(false, false, false, typeof(DisplayTestEnum), false);

            Assert.Multiple(() =>
            {
                Assert.That(attribute.EnumType, Is.EqualTo(typeof(DisplayTestEnum)));
                Assert.That(attribute.UseDisplayName, Is.True);
                Assert.That(attribute.Values, Is.EqualTo(new[] { "Yes, eligible", "No, not eligible", "Pending" }));
                Assert.That(rawAttribute.EnumType, Is.EqualTo(typeof(DisplayTestEnum)));
                Assert.That(rawAttribute.UseDisplayName, Is.False);
                Assert.That(rawAttribute.Values, Is.EqualTo(new[] { "Yes", "No", "Pending" }));
                Assert.That(attribute.AllowBlank, Is.False);
                Assert.That(attribute.ShowInputMessage, Is.False);
                Assert.That(attribute.ShowErrorMessage, Is.False);
            });
        }

        [Test]
        public void CreateSheetFromObjectWritesEnumTextDisplayNames()
        {
            MemoryStream stream = new MemoryStream();
            using (BigExcelWriter writer = new BigExcelWriter(stream))
            {
                writer.CreateSheetFromObject([new EnumTextModel()], "EnumTextDisplay");
            }

            using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = document.WorkbookPart!;
            Sheet targetSheet = workbookPart.Workbook!.Sheets!.ChildElements
                .OfType<Sheet>()
                .Single(sheet => sheet.Name?.Value == "EnumTextDisplay");
            WorksheetPart targetPart = (WorksheetPart)workbookPart.GetPartById(targetSheet.Id!.Value!);
            Row[] rows = targetPart.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
            string[] values = rows[1].Elements<Cell>()
                .Select(cell => TestHelperMethods.GetCellRealValue(cell, workbookPart))
                .ToArray();

            Assert.That(values, Is.EqualTo(new[] { "Yes, eligible", "No, not eligible", "Pending" }));
        }

        [Test]
        public void MemberConstructorResolvesEnumDisplayNames()
        {
            ExcelListValidatorAttribute attribute = new ExcelListValidatorAttribute(typeof(DisplaySource), nameof(DisplaySource.Values));
            ExcelListValidatorAttribute rawAttribute = new ExcelListValidatorAttribute(typeof(DisplaySource), nameof(DisplaySource.Values), false);

            Assert.Multiple(() =>
            {
                Assert.That(attribute.SourceType, Is.EqualTo(typeof(DisplaySource)));
                Assert.That(attribute.MemberName, Is.EqualTo(nameof(DisplaySource.Values)));
                Assert.That(attribute.UseDisplayName, Is.True);
                Assert.That(attribute.Values, Is.EqualTo(new[] { "Yes, eligible", "No, not eligible", "Pending" }));
                Assert.That(rawAttribute.UseDisplayName, Is.False);
                Assert.That(rawAttribute.Values, Is.EqualTo(new[] { "Yes", "No", "Pending" }));
            });
        }

        private enum DisplayTestEnum
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
            public DisplayTestEnum Display { get; set; } = DisplayTestEnum.Yes;

            [ExcelColumnType(CellDataType.EnumText)]
            public DisplayTestEnum NotEligible { get; set; } = DisplayTestEnum.No;

            [ExcelColumnType(CellDataType.EnumText)]
            public DisplayTestEnum PendingValue { get; set; } = DisplayTestEnum.Pending;
        }

        private static class DisplaySource
        {
            public static IEnumerable<Enum> Values
            {
                get
                {
                    return new Enum[] { DisplayTestEnum.Yes, DisplayTestEnum.No, DisplayTestEnum.Pending };
                }
            }
        }
    }
}
#endif
