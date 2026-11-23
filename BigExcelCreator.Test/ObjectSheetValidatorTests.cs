// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes;
using BigExcelCreator.Enums;
using BigExcelCreator.Extensions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BigExcelCreator.Test
{
    [TestFixture]
    internal sealed class ObjectSheetValidatorTests
    {
        [Test]
        public void CreateSheetFromObjectAppliesAllNumericOperatorsAndOperandTypes()
        {
            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream))
            {
                writer.CreateSheetFromObject([new AllValidatorsModel()], "Operators");
            }

            using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = document.WorkbookPart!;
            Sheet targetSheet = workbookPart.Workbook!.Sheets!.ChildElements
                .OfType<Sheet>()
                .Single(sheet => sheet.Name?.Value == "Operators");
            WorksheetPart targetPart = (WorksheetPart)workbookPart.GetPartById(targetSheet.Id!.Value!);
            DataValidation[] validations = targetPart.Worksheet!.ChildElements
                .OfType<DataValidations>()
                .Single()
                .ChildElements
                .OfType<DataValidation>()
                .ToArray();

            DataValidationOperatorValues[] expectedOperators =
            [
                DataValidationOperatorValues.Equal,
                DataValidationOperatorValues.NotEqual,
                DataValidationOperatorValues.GreaterThan,
                DataValidationOperatorValues.GreaterThanOrEqual,
                DataValidationOperatorValues.LessThan,
                DataValidationOperatorValues.LessThanOrEqual,
                DataValidationOperatorValues.Between,
                DataValidationOperatorValues.NotBetween,
                DataValidationOperatorValues.Between,
                DataValidationOperatorValues.Between,
                DataValidationOperatorValues.Between,
                DataValidationOperatorValues.LessThan,
                DataValidationOperatorValues.GreaterThan,
                DataValidationOperatorValues.Between,
            ];

            Assert.Multiple(() =>
            {
                Assert.That(validations.Length, Is.EqualTo(expectedOperators.Length));
                Assert.That(validations.Select(validation => validation.Operator!.Value), Is.EqualTo(expectedOperators));
                Assert.That(validations[0].Type!.Value, Is.EqualTo(DataValidationValues.Whole));
                Assert.That(validations[6].Formula1!.Text, Is.EqualTo("1"));
                Assert.That(validations[6].Formula2!.Text, Is.EqualTo("10"));
                Assert.That(validations[7].Formula1!.Text, Is.EqualTo("2"));
                Assert.That(validations[7].Formula2!.Text, Is.EqualTo("20"));
                Assert.That(validations[8].Type!.Value, Is.EqualTo(DataValidationValues.Whole));
                Assert.That(validations[8].Formula1!.Text, Is.EqualTo("3"));
                Assert.That(validations[8].Formula2!.Text, Is.EqualTo("11"));
                Assert.That(validations[9].Formula1!.Text, Is.EqualTo("4"));
                Assert.That(validations[9].Formula2!.Text, Is.EqualTo("12"));
                Assert.That(validations[10].Type!.Value, Is.EqualTo(DataValidationValues.Decimal));
                Assert.That(validations[10].Formula1!.Text, Is.EqualTo("0.5"));
                Assert.That(validations[10].Formula2!.Text, Is.EqualTo("5.5"));
                Assert.That(validations[11].Formula1!.Text, Is.EqualTo("1.25"));
                Assert.That(validations[12].Formula1!.Text, Is.EqualTo("2.5"));
                Assert.That(validations[13].Formula1!.Text, Is.EqualTo("3.75"));
                Assert.That(validations[13].Formula2!.Text, Is.EqualTo("7.25"));
            });
        }

        [Test]
        public void CreateSheetFromObjectAppliesLiteralAndFormulaListOptions()
        {
            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream))
            {
                writer.CreateSheetFromObject([new ListValidatorsModel()], "Lists");
            }

            using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = document.WorkbookPart!;
            Sheet[] sheets = workbookPart.Workbook!.Sheets!.ChildElements.OfType<Sheet>().ToArray();
            Sheet targetSheet = sheets.Single(sheet => sheet.Name?.Value == "Lists");
            WorksheetPart targetPart = (WorksheetPart)workbookPart.GetPartById(targetSheet.Id!.Value!);
            DataValidation[] validations = targetPart.Worksheet!.ChildElements
                .OfType<DataValidations>()
                .Single()
                .ChildElements
                .OfType<DataValidation>()
                .ToArray();

            Assert.Multiple(() =>
            {
                Assert.That(sheets, Has.Length.EqualTo(3));
                Assert.That(sheets.Count(sheet => sheet.State?.Value == SheetStateValues.VeryHidden), Is.EqualTo(2));
                Assert.That(validations, Has.Length.EqualTo(4));
                Assert.That(validations.Select(validation => validation.Type!.Value), Is.All.EqualTo(DataValidationValues.List));
                Assert.That(validations[0].Formula1!.Text, Does.StartWith("'").And.Contain("'!$A$1:$A$2"));
                Assert.That(validations[0].AllowBlank!.Value, Is.True);
                Assert.That(validations[1].Formula1!.Text, Does.StartWith("'").And.Contain("'!$A$1:$A$1"));
                Assert.That(validations[2].Formula1!.Text, Is.EqualTo("'Lookup'!$A$1:$A$4"));
                Assert.That(validations[2].AllowBlank!.Value, Is.True);
                Assert.That(validations[3].Formula1!.Text, Is.EqualTo("NamedCategories"));
                Assert.That(validations[3].AllowBlank!.Value, Is.False);
                Assert.That(validations[3].ShowInputMessage!.Value, Is.False);
                Assert.That(validations[3].ShowErrorMessage!.Value, Is.True);
            });

            string[] literalValues = GetListValues(workbookPart, validations[0].Formula1!.Text!);
            string[] emptyLiteralValues = GetListValues(workbookPart, validations[1].Formula1!.Text!);

            Assert.Multiple(() =>
            {
                Assert.That(literalValues, Is.EqualTo(new[] { "Alpha", string.Empty }));
                Assert.That(emptyLiteralValues, Is.EqualTo(new[] { string.Empty }));
            });
        }

        [Test]
        public void CreateSheetFromObjectCreatesSingleCellValidationRangeWhenDataIsEmpty()
        {
            MemoryStream stream = new();
            using (BigExcelWriter writer = new(stream))
            {
                writer.CreateSheetFromObject(new List<EmptyDataModel>(), "Empty", writeHeaderRow: false, columns: new List<Column>());
            }

            using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = document.WorkbookPart!;
            Sheet targetSheet = workbookPart.Workbook!.Sheets!.ChildElements
                .OfType<Sheet>()
                .Single(sheet => sheet.Name?.Value == "Empty");
            WorksheetPart targetPart = (WorksheetPart)workbookPart.GetPartById(targetSheet.Id!.Value!);
            DataValidation validation = targetPart.Worksheet!.ChildElements
                .OfType<DataValidations>()
                .Single()
                .ChildElements
                .OfType<DataValidation>()
                .Single();

            Assert.That(validation.SequenceOfReferences!.InnerText, Is.EqualTo("A1"));
        }

        [Test]
        public void ExcelListValidatorEnumConstructorResolvesDisplayNamesAndFlags()
        {
            ExcelListValidatorAttribute defaultAttribute = new ExcelListValidatorAttribute(typeof(ListValidatorTestEnum));
            ExcelListValidatorAttribute configuredAttribute = new ExcelListValidatorAttribute(false, false, false, typeof(ListValidatorTestEnum));

            string[] expectedValues = new[] { "Yes", "No", "Pending" };

            Assert.Multiple(() =>
            {
                Assert.That(defaultAttribute.EnumType, Is.EqualTo(typeof(ListValidatorTestEnum)));
                Assert.That(defaultAttribute.Values, Is.EqualTo(expectedValues));
                Assert.That(defaultAttribute.AllowBlank, Is.True);
                Assert.That(defaultAttribute.ShowInputMessage, Is.True);
                Assert.That(defaultAttribute.ShowErrorMessage, Is.True);
                Assert.That(configuredAttribute.EnumType, Is.EqualTo(typeof(ListValidatorTestEnum)));
                Assert.That(configuredAttribute.Values, Is.EqualTo(expectedValues));
                Assert.That(configuredAttribute.AllowBlank, Is.False);
                Assert.That(configuredAttribute.ShowInputMessage, Is.False);
                Assert.That(configuredAttribute.ShowErrorMessage, Is.False);
            });
        }

        [Test]
        public void ExcelListValidatorEnumConstructorRejectsNullAndNonEnumTypes()
        {
            Type nullType = null!;
            Assert.Throws<ArgumentNullException>(() => new ExcelListValidatorAttribute(nullType));
            Assert.Throws<ArgumentException>(() => new ExcelListValidatorAttribute(typeof(string)));
            Assert.Throws<ArgumentNullException>(() => new ExcelListValidatorAttribute(true, true, true, nullType));
            Assert.Throws<ArgumentException>(() => new ExcelListValidatorAttribute(true, true, true, typeof(string)));
        }

        [TestCase("PublicProperty")]
        [TestCase("InternalProperty")]
        [TestCase("PublicField")]
        [TestCase("InternalField")]
        [TestCase("PublicMethod")]
        [TestCase("InternalMethod")]
        public void ExcelListValidatorMemberConstructorResolvesPublicAndInternalStringSources(string memberName)
        {
            ExcelListValidatorAttribute defaultAttribute = new ExcelListValidatorAttribute(typeof(ListValidatorSource), memberName);
            ExcelListValidatorAttribute configuredAttribute = new ExcelListValidatorAttribute(false, false, false, typeof(ListValidatorSource), memberName);

            Assert.Multiple(() =>
            {
                Assert.That(defaultAttribute.SourceType, Is.EqualTo(typeof(ListValidatorSource)));
                Assert.That(defaultAttribute.MemberName, Is.EqualTo(memberName));
                Assert.That(defaultAttribute.Values, Is.EqualTo(new[] { "Engineering", "Finance" }));
                Assert.That(defaultAttribute.AllowBlank, Is.True);
                Assert.That(defaultAttribute.ShowInputMessage, Is.True);
                Assert.That(defaultAttribute.ShowErrorMessage, Is.True);
                Assert.That(configuredAttribute.SourceType, Is.EqualTo(typeof(ListValidatorSource)));
                Assert.That(configuredAttribute.MemberName, Is.EqualTo(memberName));
                Assert.That(configuredAttribute.Values, Is.EqualTo(new[] { "Engineering", "Finance" }));
                Assert.That(configuredAttribute.AllowBlank, Is.False);
                Assert.That(configuredAttribute.ShowInputMessage, Is.False);
                Assert.That(configuredAttribute.ShowErrorMessage, Is.False);
            });
        }

        [TestCase("PublicEnumProperty")]
        [TestCase("InternalEnumProperty")]
        [TestCase("PublicEnumMethod")]
        [TestCase("InternalEnumField")]
        public void ExcelListValidatorMemberConstructorResolvesEnumSourcesAndDisplayNames(string memberName)
        {
            ExcelListValidatorAttribute attribute = new ExcelListValidatorAttribute(typeof(ListValidatorSource), memberName);

            string[] expectedValues = new[] { "Yes", "No", "Pending" };

            Assert.Multiple(() =>
            {
                Assert.That(attribute.SourceType, Is.EqualTo(typeof(ListValidatorSource)));
                Assert.That(attribute.MemberName, Is.EqualTo(memberName));
                Assert.That(attribute.Values, Is.EqualTo(expectedValues));
            });
        }

        [Test]
        public void ExcelListValidatorMemberConstructorRejectsInvalidArgumentsAndReturnTypes()
        {
            Type nullType = null!;
            Assert.Throws<ArgumentNullException>(() => new ExcelListValidatorAttribute(nullType, "PublicProperty"));
            Assert.Throws<ArgumentException>(() => new ExcelListValidatorAttribute(typeof(ListValidatorSource), null!));
            Assert.Throws<ArgumentException>(() => new ExcelListValidatorAttribute(typeof(ListValidatorSource), string.Empty));
            Assert.Throws<ArgumentException>(() => new ExcelListValidatorAttribute(typeof(ListValidatorSource), "   "));
            Assert.Throws<ArgumentException>(() => new ExcelListValidatorAttribute(typeof(ListValidatorSource), "MissingMember"));
            Assert.Throws<InvalidOperationException>(() => new ExcelListValidatorAttribute(typeof(ListValidatorSource), "NullProperty"));
            Assert.Throws<InvalidOperationException>(() => new ExcelListValidatorAttribute(typeof(ListValidatorSource), "StringProperty"));
            Assert.Throws<InvalidOperationException>(() => new ExcelListValidatorAttribute(typeof(ListValidatorSource), "IntegerArrayField"));
        }

        [Test]
        public void FormulaListValidatorRejectsNullFormula()
        {
            Assert.Throws<ArgumentNullException>(() => new ExcelListFormulaValidatorAttribute(null!));
        }

        [Test]
        public void ValidatorOperatorRejectsUndefinedValues()
        {
            Assert.Throws<ArgumentOutOfRangeException>(
                () => ((ValidatorOperator)int.MaxValue).ToDataValidationOperatorValues());
        }

        [Test]
        public void ValidatorListSheetNameHandlesShortBlankAndInvalidParts()
        {
            string shortName = "Orders".CreateValidatorListSheetName("Category");
            string blankPartsName = " ".CreateValidatorListSheetName("?");
            string invalidCharactersName = "Sheet/Name".CreateValidatorListSheetName("Column[Name]");

            Assert.Multiple(() =>
            {
                Assert.That(shortName, Does.StartWith("OrdersCategory"));
                Assert.That(shortName, Has.Length.EqualTo(26));
                Assert.That(blankPartsName, Does.StartWith("ValidatioValidation"));
                Assert.That(invalidCharactersName, Does.StartWith("SheetNameColumnName"));
                Assert.That(invalidCharactersName, Does.Not.Contain("/"));
                Assert.That(invalidCharactersName, Does.Not.Contain("["));
            });
        }

        private static string[] GetListValues(WorkbookPart workbookPart, string formula)
        {
            string sheetName = formula.Split('!')[0].Trim('\'');
            Sheet sheet = workbookPart.Workbook!.Sheets!.ChildElements
                .OfType<Sheet>()
                .Single(item => item.Name?.Value == sheetName);
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);

            return worksheetPart.Worksheet!.ChildElements.OfType<SheetData>().Single()
                .ChildElements.OfType<Row>()
                .SelectMany(row => row.ChildElements.OfType<Cell>())
                .Select(cell => cell.CellValue?.Text ?? string.Empty)
                .ToArray();
        }

        private sealed class AllValidatorsModel
        {
            [ExcelIntegerValidator(1, ValidatorOperator.Equal)]
            public int IntegerEqual { get; set; }

            [ExcelIntegerValidator(1U, ValidatorOperator.NotEqual)]
            public uint IntegerNotEqual { get; set; }

            [ExcelIntegerValidator(1L, ValidatorOperator.GreaterThan)]
            public long IntegerGreaterThan { get; set; }

            [ExcelIntegerValidator(1UL, ValidatorOperator.GreaterThanOrEqual)]
            public ulong IntegerGreaterThanOrEqual { get; set; }

            [ExcelIntegerValidator(1, ValidatorOperator.LessThan)]
            public int IntegerLessThan { get; set; }

            [ExcelIntegerValidator(1U, ValidatorOperator.LessThanOrEqual)]
            public uint IntegerLessThanOrEqual { get; set; }

            [ExcelIntegerValidator(1L, ValidatorOperator.Between, secondOperand: 10L)]
            public long IntegerBetween { get; set; }

            [ExcelIntegerValidator(2UL, ValidatorOperator.NotBetween, secondOperand: 20UL)]
            public ulong IntegerNotBetween { get; set; }

            [ExcelIntegerValidator(3, ValidatorOperator.Between, secondOperand: 11)]
            public int IntegerBetweenInt { get; set; }

            [ExcelIntegerValidator(4U, ValidatorOperator.Between, secondOperand: 12U)]
            public uint UnsignedIntegerBetween { get; set; }

            [ExcelDecimalValidator(0.5d, ValidatorOperator.Between, secondOperand: 5.5d)]
            public decimal DecimalBetween { get; set; }

            [ExcelDecimalValidator(1.25f, ValidatorOperator.LessThan)]
            public float FloatLessThan { get; set; }

            [ExcelDecimalValidator(2.5d, ValidatorOperator.GreaterThan)]
            public double DoubleGreaterThan { get; set; }

            [ExcelDecimalValidator(3.75f, ValidatorOperator.Between, secondOperand: 7.25f)]
            public float FloatBetween { get; set; }
        }

        private sealed class ListValidatorsModel
        {
            [ExcelListValidator("Alpha", null!)]
            public string LiteralList { get; set; } = string.Empty;

            [ExcelListValidator()]
            public string EmptyLiteralList { get; set; } = string.Empty;

            [ExcelListFormulaValidator("'Lookup'!$A$1:$A$4")]
            public string DefaultFormulaList { get; set; } = string.Empty;

            [ExcelListFormulaValidator("NamedCategories", false, false, true)]
            public string ConfiguredFormulaList { get; set; } = string.Empty;
        }

        internal enum ListValidatorTestEnum
        {
            Yes,

            No,

            Pending
        }

        private sealed class ListValidatorSource
        {
            public static IEnumerable<string> PublicProperty
            {
                get
                {
                    return new[] { "Engineering", "Finance", "Engineering" }.Distinct();
                }
            }

            internal static IEnumerable<string> InternalProperty
            {
                get
                {
                    return new[] { "Engineering", "Finance", "Engineering" }.Distinct();
                }
            }

            public static IEnumerable<string> PublicField = new[] { "Engineering", "Finance", "Engineering" }.Distinct();

            internal static IEnumerable<string> InternalField = new[] { "Engineering", "Finance", "Engineering" }.Distinct();

            public static IEnumerable<string> PublicMethod()
            {
                return new[] { "Engineering", "Finance", "Engineering" }.Distinct();
            }

            internal static IEnumerable<string> InternalMethod()
            {
                return new[] { "Engineering", "Finance", "Engineering" }.Distinct();
            }

            public static IEnumerable<Enum> PublicEnumProperty
            {
                get
                {
                    return new Enum[] { ListValidatorTestEnum.Yes, ListValidatorTestEnum.No, ListValidatorTestEnum.Pending };
                }
            }

            internal static IEnumerable<Enum> InternalEnumProperty
            {
                get
                {
                    return new Enum[] { ListValidatorTestEnum.Yes, ListValidatorTestEnum.No, ListValidatorTestEnum.Pending };
                }
            }

            public static IEnumerable<Enum> PublicEnumMethod()
            {
                return new Enum[] { ListValidatorTestEnum.Yes, ListValidatorTestEnum.No, ListValidatorTestEnum.Pending };
            }

            internal static IEnumerable<Enum> InternalEnumField = new Enum[] { ListValidatorTestEnum.Yes, ListValidatorTestEnum.No, ListValidatorTestEnum.Pending };

            public static IEnumerable<string> NullProperty
            {
                get
                {
                    return null!;
                }
            }

            public static string StringProperty
            {
                get
                {
                    return "Engineering";
                }
            }

            public static int[] IntegerArrayField = new[] { 1, 2, 3 };
        }

        private sealed class EmptyDataModel
        {
            [ExcelListFormulaValidator("AvailableValues")]
            public string Value { get; set; } = string.Empty;
        }
    }
}
