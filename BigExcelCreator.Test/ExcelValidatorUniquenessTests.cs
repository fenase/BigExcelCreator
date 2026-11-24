// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes;
using BigExcelCreator.Enums;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BigExcelCreator.Test
{
    [TestFixture]
    internal sealed class ExcelValidatorUniquenessTests
    {
        [Test]
        public void SingleValidatorContinuesToApply()
        {
            using MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream))
            {
                writer.CreateSheetFromObject([new SingleValidatorModel()], "SingleValidator");
            }

            using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(
                document.WorkbookPart.Workbook!.Sheets!.Elements<Sheet>().Single().Id!.Value!);
            DataValidation validation = worksheetPart.Worksheet!.Elements<DataValidations>().Single()
                .Elements<DataValidation>().Single();

            Assert.Multiple(() =>
            {
                Assert.That(validation.Type!.Value, Is.EqualTo(DataValidationValues.Whole));
                Assert.That(validation.Formula1!.Text, Is.EqualTo("1"));
            });
        }

        [Test]
        public void DifferentValidatorsOnOnePropertyThrowWithModelPropertyAndValidatorTypes()
        {
            using MemoryStream stream = new();
            using BigExcelWriter writer = new(stream);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => writer.CreateSheetFromObject([new ConflictingValidatorsModel()], "Conflicting"));

            Assert.Multiple(() =>
            {
                Assert.That(exception!.Message, Does.Contain(typeof(ConflictingValidatorsModel).FullName!));
                Assert.That(exception.Message, Does.Contain(nameof(ConflictingValidatorsModel.Value)));
                Assert.That(exception.Message, Does.Contain(typeof(ExcelIntegerValidatorAttribute).FullName!));
                Assert.That(exception.Message, Does.Contain(typeof(ExcelDecimalValidatorAttribute).FullName!));
            });
        }

        private sealed class SingleValidatorModel
        {
            [ExcelIntegerValidator(1, ValidatorOperator.Equal)]
            public int Value { get; set; } = 1;
        }

        private sealed class ConflictingValidatorsModel
        {
            [ExcelIntegerValidator(1, ValidatorOperator.Equal)]
            [ExcelDecimalValidator(2.0, ValidatorOperator.Equal)]
            public decimal Value { get; set; } = 1;
        }
    }
}
