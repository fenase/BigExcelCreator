using BigExcelCreator.ClassAttributes;
using BigExcelCreator.Enums;

namespace BigExcelCreator.Test.Classes
{
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
        [ExcelColumnOrder(8)]
        [ExcelColumnType(CellDataType.Text)]
        [ExcelColumnWidth(20)]
        public DateTime CreatedAt { get; set; }

        [ExcelColumnName("Monthly Salary")]
        [ExcelColumnOrder(4)]
        [ExcelColumnType(CellDataType.Number)]
        [ExcelColumnWidth(20)]
        public decimal Salary { get; set; }

        [ExcelColumnName("Sale Amount")]
        [ExcelColumnOrder(5)]
        [ExcelColumnType(CellDataType.Number)]
        [ExcelColumnWidth(20)]
        public decimal Sale { get; set; }

        [ExcelColumnName("Additional Notes")]
        [ExcelColumnOrder(6)]
        [ExcelColumnType(CellDataType.Text)]
        [ExcelColumnWidth(40)]
        string? Notes { get; set; }

        [ExcelColumnHidden]
        public string Secret { get; set; } = "";

        [ExcelColumnName("Public Information")]
        [ExcelColumnOrder(7)]
        [ExcelColumnType(CellDataType.Text)]
        [ExcelColumnWidth(40)]
        public string PublicInfo { get; set; } = "";

        [ExcelIgnore]
        public int TopSecretNumber { get; set; }

        internal static List<ExampleModel> GetTestData() =>
            new()
            {
                new ExampleModel
                {
                    Id = 1,
                    Name = "John Doe",
                    Position = "Software Engineer",
                    Description = "Responsible for developing software solutions.",
                    CreatedAt = DateTime.Now.AddMonths(-6),
                    Salary = 6000.50m,
                    Sale = 15000.75m,
                    Notes = "Excellent performance.",
                    Secret = "Loves pizza",
                    PublicInfo = "Enjoys hiking.",
                    TopSecretNumber= 42,
                },
                new ExampleModel
                {
                    Id = 2,
                    Name = "Jane Smith",
                    Position = "Project Manager",
                    Description = "Oversees project development and delivery.",
                    CreatedAt = DateTime.Now.AddYears(-1),
                    Salary = 8000.00m,
                    Sale = 20000.00m,
                    Notes = "Strong leadership skills.",
                    Secret = "Collects stamps",
                    PublicInfo = "Volunteers at local shelter.",
                    TopSecretNumber= 7,
                },
                new ExampleModel
                {
                    Id = 3,
                    Name = "Alice Johnson",
                    Position = "UX Designer",
                    Description = "Designs user-friendly interfaces.",
                    CreatedAt = DateTime.Now.AddMonths(-3),
                    Salary = 5500.25m,
                    Sale = 12000.50m,
                    Notes = "Creative thinker.",
                    Secret = "Plays the violin",
                    PublicInfo = "Loves painting.",
                    TopSecretNumber= 13,
                },
                new ExampleModel
                {
                    Id = 4,
                    Name = "Bob Brown",
                    Position = "Data Analyst",
                    Description = "Analyzes data to support business decisions.",
                    CreatedAt = DateTime.Now.AddMonths(-9),
                    Salary = 6200.75m,
                    Sale = 18000.00m,
                    Notes = "Detail-oriented.",
                    Secret = "Enjoys skydiving",
                    PublicInfo = "Active in community sports.",
                    TopSecretNumber= 99,
                },
                new ExampleModel
                {
                    Id = 5,
                    Name = "Eve Davis",
                    Position = "Marketing Specialist",
                    Description = "Develops marketing strategies and campaigns.",
                    CreatedAt = DateTime.Now.AddYears(-2),
                    Salary = 7000.00m,
                    Sale = 22000.25m,
                    Notes = "Excellent communication skills.",
                    Secret = "Writes poetry",
                    PublicInfo = "Travels frequently.",
                    TopSecretNumber= 56,
                }
            };
    }
}
