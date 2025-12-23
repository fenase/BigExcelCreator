using BigExcelCreator.ClassAttributes;
using BigExcelCreator.Enums;

namespace BigExcelCreator.Test.Classes
{
    [ExcelHeaderStyleFormat(6)]
    internal class ExampleModel
    {
        [ExcelColumnName("Identifier")]
        [ExcelColumnOrder(0)]
        [ExcelColumnType(CellDataType.Number)]
        [ExcelColumnWidth(15)]
        public int Id { get; set; }

        [ExcelColumnName("Full Name")]
        [ExcelColumnOrder(1)]
        [ExcelColumnType(CellDataType.Text)]
        [ExcelColumnWidth(30)]
        public string Name { get; set; } = "";

        [ExcelColumnName("Job Position")]
        [ExcelColumnOrder(2)]
        [ExcelColumnType(CellDataType.Text)]
        [ExcelColumnWidth(25)]
        public string Position { get; set; } = "";

        [ExcelColumnName("Job Description")]
        [ExcelColumnOrder(3)]
        [ExcelColumnType(CellDataType.Text)]
        [ExcelColumnWidth(50)]
        public string Description { get; set; } = "";

        [ExcelColumnName("Creation Date")]
        [ExcelColumnOrder(10)]
        [ExcelColumnType(CellDataType.Text)]
        [ExcelColumnWidth(20)]
        public DateTime CreatedAt { get; set; }

        [ExcelColumnName("Monthly Salary")]
        [ExcelColumnOrder(4)]
        [ExcelColumnType(CellDataType.Number)]
        [ExcelColumnWidth(20)]
        public decimal Salary { get; set; }

        [ExcelColumnName("Monthly Bonus")]
        [ExcelColumnOrder(5)]
        [ExcelColumnType(CellDataType.Number)]
        [ExcelColumnWidth(20)]
        [ExcelStyleFormat(8)]
        public double? Bonus { get; set; }

        [ExcelColumnName("Monthly Net Income")]
        [ExcelColumnOrder(6)]
        [ExcelColumnType(CellDataType.Formula)]
        [ExcelColumnWidth(20)]
        [ExcelStyleFormat(8, StylingPriority.Data)]
        public string? NetIncome { get; set; } = "";

        [ExcelColumnName("Sale Amount")]
        [ExcelColumnOrder(7)]
        [ExcelColumnType(CellDataType.Number)]
        [ExcelColumnWidth(20)]
        public float? Sale { get; set; }

        [ExcelColumnName("Additional Notes")]
        [ExcelColumnOrder(8)]
        [ExcelColumnType(CellDataType.Text)]
        [ExcelColumnWidth(40)]
        string? Notes { get; set; }

        [ExcelColumnHidden]
        public string Secret { get; set; } = "";

        [ExcelColumnName("Public Information")]
        [ExcelColumnOrder(9)]
        [ExcelColumnType(CellDataType.Text)]
        [ExcelColumnWidth(40)]
        public string PublicInfo { get; set; } = "";

        [ExcelIgnore]
        public int TopSecretNumber { get; set; }
    }
}
