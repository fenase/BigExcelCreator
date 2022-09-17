// See https://aka.ms/new-console-template for more information

using BigExcelCreator;
using DocumentFormat.OpenXml.Spreadsheet;

int attemps = 0;
string fullpath;
do
{
    string path = Path.GetTempPath() + @"excelTest\";
    Directory.CreateDirectory(path);
    string name = DateTime.Now.ToString("yyyyMMddHHmmssff") + ".xlsx";
    fullpath = Path.Combine(path, name);
} while (attemps < 10 && File.Exists(fullpath));

Console.WriteLine(fullpath);

using BigExcelWritter excel = new(fullpath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);

excel.CreateAndOpenSheet("S1", sheetState: SheetStateValues.Visible);


excel.WriteTextRow(new List<string> { "A1", "B1", "C1", "D1", "E1" });
excel.WriteTextRow(new List<string> { "A2", "B2", "C2", "D2", "E2" });
excel.WriteTextRow(new List<string> { "A3", "B3", "C3", "D3", "E3" });


excel.Comment("test A1", "A1", "Me");
excel.Comment("test A3", "A3", "you");
excel.Comment("test B2", "B2");
excel.Comment("test E2", "E2", "unknown");

excel.CloseSheet();

excel.CreateAndOpenSheet("S2");

excel.WriteTextRow(new List<string> { "A1", "B1", "C1", "D1", "E1" });
excel.WriteTextRow(new List<string> { "A2", "B2", "C2", "D2", "E2" });
excel.WriteTextRow(new List<string> { "A3", "B3", "C3", "D3", "E3" });

excel.Comment("test A1 another sheet", "A1", "Me");
excel.Comment("test E3 another sheet", "E3", "you too");
excel.Comment("test B2 another sheet", "B2");

excel.BeginRow();
excel.WriteTextCell("new cell??");

excel.Comment("comment while writing row", "C4", "me");

excel.EndRow();


excel.CloseSheet();


