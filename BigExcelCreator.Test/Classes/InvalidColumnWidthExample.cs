using BigExcelCreator.ClassAttributes;

namespace BigExcelCreator.Test.Classes
{
    internal class InvalidColumnWidthExample
    {
        [ExcelColumnName("First Column")]
        [ExcelColumnWidth(-10)]
        public int FirstColumn { get; set; }
    }
}
