// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

#if NET48_OR_GREATER || NET6_0_OR_GREATER || NET8_0_OR_GREATER || NET10_0_OR_GREATER
using BigExcelCreator.ClassAttributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace BigExcelCreator.Test
{
    [TestFixture]
    internal sealed class ExcelListValidatorAttributeDisplayTests
    {
        [Test]
        public void EnumConstructorResolvesDisplayNamesAndFallbacks()
        {
            ExcelListValidatorAttribute attribute = new ExcelListValidatorAttribute(typeof(DisplayTestEnum));

            Assert.Multiple(() =>
            {
                Assert.That(attribute.EnumType, Is.EqualTo(typeof(DisplayTestEnum)));
                Assert.That(attribute.Values, Is.EqualTo(new[] { "Yes, eligible", "No, not eligible", "Pending" }));
            });
        }

        [Test]
        public void EnumConstructorWithDisplayOptionsUsesResolvedNamesAndFlags()
        {
            ExcelListValidatorAttribute attribute = new ExcelListValidatorAttribute(false, false, false, typeof(DisplayTestEnum));

            Assert.Multiple(() =>
            {
                Assert.That(attribute.EnumType, Is.EqualTo(typeof(DisplayTestEnum)));
                Assert.That(attribute.Values, Is.EqualTo(new[] { "Yes, eligible", "No, not eligible", "Pending" }));
                Assert.That(attribute.AllowBlank, Is.False);
                Assert.That(attribute.ShowInputMessage, Is.False);
                Assert.That(attribute.ShowErrorMessage, Is.False);
            });
        }

        [Test]
        public void MemberConstructorResolvesEnumDisplayNames()
        {
            ExcelListValidatorAttribute attribute = new ExcelListValidatorAttribute(typeof(DisplaySource), nameof(DisplaySource.Values));

            Assert.Multiple(() =>
            {
                Assert.That(attribute.SourceType, Is.EqualTo(typeof(DisplaySource)));
                Assert.That(attribute.MemberName, Is.EqualTo(nameof(DisplaySource.Values)));
                Assert.That(attribute.Values, Is.EqualTo(new[] { "Yes, eligible", "No, not eligible", "Pending" }));
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
