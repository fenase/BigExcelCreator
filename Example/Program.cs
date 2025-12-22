// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// See https://aka.ms/new-console-template for more information

using BigExcelCreator;
using BigExcelCreator.Styles;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Example;

int attempts = 0;
string fullPath;
do
{
    string path = Path.GetTempPath() + @"excelTest\";
    Directory.CreateDirectory(path);
    string name = DateTime.Now.ToString("yyyyMMddHHmmssff") + ".xlsx";
    fullPath = Path.Combine(path, name);
} while (attempts++ < 10 && File.Exists(fullPath));

Console.WriteLine(fullPath);

var columns = new List<Column>
{
    new Column{ CustomWidth=true, Width=15, Hidden = true },
    new Column{ Width=15, Hidden = false },
    new Column{ CustomWidth=true, Width=19},
    new Column{ CustomWidth=true, Width=5, Hidden = true },
    new Column{ Hidden = true },
    new Column{ Hidden = false },
};

StyleList styleList = new();
Font italic = new(new Italic());
Font bold = new(new Bold());
Font boldItalic = new(new Bold(), new Italic());
styleList.NewStyle(italic, null, null, null, "italic default");
styleList.NewStyle(bold, null, null, null, "bold default");
styleList.NewStyle(boldItalic, null, null, null, "bold italic default");

Alignment center = new() { Horizontal = HorizontalAlignmentValues.Center };

styleList.NewStyle(italic, null, null, null, center, "italic center");
styleList.NewStyle(bold, null, null, null, center, "bold center");
styleList.NewStyle(boldItalic, null, null, null, center, "bold italic center");
Fill yellowFill = new Fill(new[]{
                        new PatternFill(new[]{
                            new ForegroundColor { Rgb = new HexBinaryValue { Value = "FFFF00" } } }
                        )
                        { PatternType = PatternValues.Solid } });
styleList.NewStyle(null, yellowFill, null, null, "YELLOW");

styleList.NewDifferentialStyle("RED", font: new Font(new[] { new Color { Rgb = new HexBinaryValue { Value = "FF0000" } } }));

Fill greenFill = new Fill(new[]{
                        new PatternFill(new[]{
                            new BackgroundColor { Rgb = new HexBinaryValue { Value = "00FF00" } } })
                        { PatternType = PatternValues.Solid } });

styleList.NewDifferentialStyle("GREENBKG", fill: greenFill);

using BigExcelWriter excel = new(fullPath, styleList.GetStylesheet());

excel.CreateAndOpenSheet("S1", columns: columns, sheetState: SheetStateValues.Visible);

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
excel.WriteTextRow(new List<string> { "A2", "B2", "C2", "D2", "E2" }, hidden: true);
excel.WriteTextRow(new List<string> { "A3", "B3", "C3", "D3", "E3" });
excel.WriteNumberRow(new List<float> { 548, 1872, 14663, 1145, 1146 });

excel.Comment("test A1 another sheet", "A1", "Me");
excel.Comment("test E3 another sheet", "E3", "you too");
excel.Comment("test B2 another sheet", "B2");

excel.BeginRow();
excel.WriteTextCell("Formulas:");
excel.WriteFormulaCell("SUM(A4:E4)");

excel.Comment("comment while writing row", "C4", "me");

excel.AddAutofilter("a1:e1");

excel.EndRow();

excel.PrintGridLinesInCurrentSheet = true;
excel.PrintRowAndColumnHeadingsInCurrentSheet = true;
excel.ShowGridLinesInCurrentSheet = false;

excel.CloseSheet();

excel.CreateAndOpenSheet("ReadMe example");
excel.BeginRow();
excel.WriteTextCell("Cell content");
excel.WriteNumberCell(123); // write as number. This allows to use formulas.
excel.WriteNumberCell(456);
excel.WriteFormulaCell("SUM(B1:C1)");
excel.EndRow();
excel.BeginRow(true);
excel.WriteTextCell("This row id hidden");
excel.EndRow();
excel.CloseSheet();

excel.CreateAndOpenSheet("format");
excel.BeginRow();
excel.WriteTextCell("this is in italic", styleList.GetIndexByName("italic default"));
excel.WriteTextCell("this is bold", styleList.GetIndexByName("bold default"));
excel.WriteTextCell("this is bold and italic", styleList.GetIndexByName("bold italic default"));
excel.EndRow();
excel.BeginRow();
excel.WriteTextCell("this is in italic (centered)", styleList.GetIndexByName("italic center"));
excel.WriteTextCell("this is bold (centered)", styleList.GetIndexByName("bold center"));
excel.WriteTextCell("this is bold and italic (centered)", styleList.GetIndexByName("bold italic center"));
excel.EndRow();
excel.CloseSheet();

excel.CreateAndOpenSheet("conditional");
for (int i = 0; i < 10; i++)
{
    excel.WriteNumberRow(new List<float> { i }, styleList.GetIndexByName("YELLOW"));
    excel.WriteNumberRow(new List<float> { i }, styleList.GetIndexByName("YELLOW"));
}

const string conditionalFormattingRange = "A1:A20";

excel.AddConditionalFormattingFormula(conditionalFormattingRange, "A1<5", styleList.GetIndexDifferentialByName("RED"));
excel.AddConditionalFormattingFormula(conditionalFormattingRange, "A1>5", styleList.GetIndexDifferentialByName("GREENBKG"));
excel.AddConditionalFormattingDuplicatedValues(conditionalFormattingRange, styleList.GetIndexDifferentialByName("RED"));
excel.AddConditionalFormattingCellIs(conditionalFormattingRange, ConditionalFormattingOperatorValues.LessThan, "5", styleList.GetIndexDifferentialByName("RED"));
excel.AddConditionalFormattingCellIs(conditionalFormattingRange, ConditionalFormattingOperatorValues.Between, "3", styleList.GetIndexDifferentialByName("RED"), "7");

excel.AddIntegerValidator("B1:B10", 2, DataValidationOperatorValues.Between, secondOperand: 8);
excel.AddDecimalValidator("C1:C10", 0d, DataValidationOperatorValues.Between, secondOperand: 3.14159265359d);

excel.CloseSheet();

excel.CreateAndOpenSheet("merged cells");

excel.MergeCells("1");
excel.MergeCells("a2:c2");
excel.MergeCells("a3:a5");
excel.MergeCells("c3:d5");

excel.CloseSheet();



excel.CreateSheetFromObject(ExampleModel.GetTestData(), "From objects");


