// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.CommentsManager;
using BigExcelCreator.Extensions;
using BigExcelCreator.Ranges;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
using System.Threading.Tasks;
#endif

[assembly: CLSCompliant(true)]
[assembly: InternalsVisibleTo("Test")]
[assembly: InternalsVisibleTo("Test35")]
[assembly: InternalsVisibleTo("Test48")]

namespace BigExcelCreator
{
    /// <summary>
    /// This class writes Excel files directly using OpenXML SAX.
    /// Useful when trying to write tens of thousands of rows.
    /// <see cref="https://www.nuget.org/packages/BigExcelCreator/#readme-body-tab">NuGet</see>
    /// <seealso cref="https://github.com/fenase/BigExcelCreator">Source</seealso>
    /// </summary>
    public class BigExcelWriter : IDisposable
    {
        #region props
        public string Path { get; }
        public Stream Stream { get; }
        private SavingTo SavingTo { get; }

        public SpreadsheetDocumentType SpreadsheetDocumentType { get; private set; }

        public SpreadsheetDocument Document { get; }

        public bool SkipCellWhenEmpty { get; set; }

        private bool sheetOpen;
        private string currentSheetName = "";
        private uint currentSheetId = 1;
        private SheetStateValues currentSheetState = SheetStateValues.Visible;
        private bool open = true;
        private int lastRowWritten;
        private bool rowOpen;
        private int columnNum = 1;
        private int maxColumnNum = 1;

        private readonly List<Sheet> sheets = new();

        private DataValidations sheetDataValidations;

        private OpenXmlWriter workSheetPartWriter;

        private readonly List<string> SharedStringsList = new();

        private WorksheetPart workSheetPart;

        private CommentManager commentManager;

        private AutoFilter SheetAutofilter;

        private SharedStringTablePart SharedStringTablePart;

        private readonly List<ConditionalFormatting> conditionalFormattingList = new();

        private readonly List<CellRange> SheetMergedCells = new();

#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
        private readonly List<Task> DocumentTasks = new();
        private readonly List<Task> SheetTasks = new();
#endif
        #endregion

        #region ctor
        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType)
        : this(stream, spreadsheetDocumentType, false) { }

        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
        : this(stream, spreadsheetDocumentType, false, stylesheet) { }

        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty)
            : this(stream, spreadsheetDocumentType, skipCellWhenEmpty, new Stylesheet()) { }

        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType)
        : this(path, spreadsheetDocumentType, false) { }

        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
        : this(path, spreadsheetDocumentType, false, stylesheet) { }

        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty)
            : this(path, spreadsheetDocumentType, skipCellWhenEmpty, new Stylesheet()) { }

        public BigExcelWriter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            Path = path;
            SavingTo = SavingTo.file;
            Document = SpreadsheetDocument.Create(Path, spreadsheetDocumentType);
            CtorHelper(spreadsheetDocumentType, skipCellWhenEmpty, stylesheet);
        }

        public BigExcelWriter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            Stream = stream;
            SavingTo = SavingTo.stream;
            Document = SpreadsheetDocument.Create(Stream, spreadsheetDocumentType);
            CtorHelper(spreadsheetDocumentType, skipCellWhenEmpty, stylesheet);
        }

        private void CtorHelper(SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            SpreadsheetDocumentType = spreadsheetDocumentType;
            WorkbookPart workbookPart = Document.AddWorkbookPart();

            if (workbookPart.WorkbookStylesPart == null)
            {
                WorkbookStylesPart wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();
                // add styles to sheet
                wbsp.Stylesheet = stylesheet;
                wbsp.Stylesheet.Save();
            }

            SharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();

            SkipCellWhenEmpty = skipCellWhenEmpty;
        }
        #endregion

        public void CreateAndOpenSheet(string name, IList<Column> columns = null,
                                       SheetStateValues sheetState = SheetStateValues.Visible)
        {
            if (!sheetOpen)
            {
                workSheetPart = Document.WorkbookPart.AddNewPart<WorksheetPart>();
                workSheetPartWriter = OpenXmlWriter.Create(workSheetPart);
                currentSheetName = name;
                workSheetPartWriter.WriteStartElement(new Worksheet());

                if (columns?.Count > 0)
                {
                    workSheetPartWriter.WriteStartElement(new Columns());
                    int indiceColumna = 1;
                    foreach (Column column in columns)
                    {
                        List<OpenXmlAttribute> atributosColumna = new()
                        {
                            new OpenXmlAttribute("min", null, indiceColumna.ToString(CultureInfo.InvariantCulture)),
                            new OpenXmlAttribute("max", null, indiceColumna.ToString(CultureInfo.InvariantCulture)),
                            new OpenXmlAttribute("width", null, (column.Width ?? 11).ToString()),
                            new OpenXmlAttribute("customWidth", null, (column.CustomWidth ?? true).ToString()),
                            new OpenXmlAttribute("hidden", null, (column.Hidden ?? false).ToString()),
                        };

                        workSheetPartWriter.WriteStartElement(new Column(), atributosColumna);
                        workSheetPartWriter.WriteEndElement();
                        ++indiceColumna;
                    }
                    workSheetPartWriter.WriteEndElement();
                }

                workSheetPartWriter.WriteStartElement(new SheetData());
                sheetOpen = true;
                currentSheetState = sheetState;
            }
            else
            {
                throw new InvalidOperationException("Sheet is already open");
            }
        }

        public void CloseSheet()
        {
            if (sheetOpen)
            {
                // write the end SheetData element
                workSheetPartWriter.WriteEndElement();

                WriteValidations();

                WriteFilters();

                WriteConditionalFormatting();

                WriteMergedCells();

                // write the end Worksheet element
                workSheetPartWriter.WriteEndElement();

                workSheetPartWriter.Close();
                workSheetPartWriter = null;

                if (commentManager != null)
                {
                    commentManager.SaveComments(workSheetPart);
                }

                sheets.Add(new Sheet()
                {
                    Name = currentSheetName,
                    SheetId = currentSheetId++,
                    Id = Document.WorkbookPart.GetIdOfPart(workSheetPart),
                    State = currentSheetState,
                });

                currentSheetName = "";
                workSheetPart.Worksheet.SheetDimension = new SheetDimension() { Reference = $"A1:{Helpers.GetColumnName(maxColumnNum)}{Math.Max(1, lastRowWritten)}" };



#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
                Task.WaitAll(SheetTasks.ToArray());
                SheetTasks.Clear();
#endif
                sheetOpen = false;
                workSheetPart = null;
                commentManager = null;
                lastRowWritten = 0;
            }
            else
            {
                throw new InvalidOperationException("There is no open sheet");
            }
        }


        public void BeginRow(int rownum)
        {
            BeginRow(rownum, false);
        }

        public void BeginRow(int rownum, bool hidden)
        {
            if (sheetOpen && !rowOpen)
            {
                if (rownum > lastRowWritten)
                {
                    lastRowWritten = rownum;
                    //create a new list of attributes
                    List<OpenXmlAttribute> attributes = new()
                    {
                        // add the row index attribute to the list
                        new OpenXmlAttribute("r", null, lastRowWritten.ToString(CultureInfo.InvariantCulture)),
                        
                        // Hide row if requested
                        new OpenXmlAttribute("hidden", null, hidden ? "1" : "0"),
                    };

                    //write the row start element with the row index attribute
                    workSheetPartWriter.WriteStartElement(new Row(), attributes);
                    rowOpen = true;
                }
                else
                {
                    throw new InvalidOperationException("Out of order row writing is not allowed");
                }
            }
            else
            {
                throw new InvalidOperationException("A row is already open. Use EndRow to close it.");
            }
        }

        public void BeginRow()
        {
            BeginRow(false);
        }

        public void BeginRow(bool hidden)
        {
            BeginRow(lastRowWritten + 1, hidden);
        }

        public void EndRow()
        {
            if (rowOpen)
            {
                // write the end row element
                workSheetPartWriter.WriteEndElement();
                maxColumnNum = Math.Max(columnNum - 1, maxColumnNum);
                columnNum = 1;
                rowOpen = false;
            }
            else
            {
                throw new InvalidOperationException("There is no row open");
            }
        }

        public void WriteTextCell(string text, int format = 0, bool useSharedStrings = false)
        {
            if (format < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(format));
            }

            if (rowOpen)
            {
                if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(text)))
                {
                    List<OpenXmlAttribute> attributes;
                    if (useSharedStrings)
                    {
                        string ssPos = AddTextToSharedStringsTable(text).ToString(CultureInfo.InvariantCulture);
                        attributes = new()
                        {
                            new OpenXmlAttribute("t", null, "s"),
                            new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture,"{0}{1}", Helpers.GetColumnName(columnNum), lastRowWritten)),
                            //styles
                            new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                        };
                        //write the cell start element with the type and reference attributes
                        workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                        //write the cell value
                        workSheetPartWriter.WriteElement(new CellValue(ssPos));
                    }
                    else
                    {
                        //reset the list of attributes
                        attributes = new()
                        {
                            // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                            new OpenXmlAttribute("t", null, "str"),
                            //add the cell reference attribute
                            new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture,"{0}{1}", Helpers.GetColumnName(columnNum), lastRowWritten)),
                            //styles
                            new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                        };
                        //write the cell start element with the type and reference attributes
                        workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                        //write the cell value
                        workSheetPartWriter.WriteElement(new CellValue(text));
                    }

                    // write the end cell element
                    workSheetPartWriter.WriteEndElement();

                }
                columnNum++;
            }
            else
            {
                throw new InvalidOperationException("There is no active row");
            }
        }

        public void WriteNumberCell(float number, int format = 0)
        {
            if (format < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(format));
            }

            if (rowOpen)
            {
                //reset the list of attributes
                List<OpenXmlAttribute> attributes = new()
                {
                    //add the cell reference attribute
                    new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture,"{0}{1}", Helpers.GetColumnName(columnNum), lastRowWritten)),
                    //styles
                    new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                };

                //write the cell start element with the type and reference attributes
                workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                //write the cell value
                workSheetPartWriter.WriteElement(new CellValue(number.ToString(CultureInfo.InvariantCulture)));

                // write the end cell element
                workSheetPartWriter.WriteEndElement();

                columnNum++;
            }
            else
            {
                throw new InvalidOperationException("There is no active row");
            }
        }

        public void WriteFormulaCell(string formula, int format = 0)
        {
            if (format < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(format));
            }

            if (rowOpen)
            {
                if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(formula)))
                {
                    //reset the list of attributes
                    List<OpenXmlAttribute> attributes = new()
                    {
                        //add the cell reference attribute
                        new OpenXmlAttribute("r", "", string.Format(CultureInfo.InvariantCulture,"{0}{1}", Helpers.GetColumnName(columnNum), lastRowWritten)),
                        //styles
                        new OpenXmlAttribute("s", null, format.ToString(CultureInfo.InvariantCulture))
                    };

                    //write the cell start element with the type and reference attributes
                    workSheetPartWriter.WriteStartElement(new Cell(), attributes);
                    //write the cell value
                    workSheetPartWriter.WriteElement(new CellFormula(formula?.ToUpperInvariant()));

                    // write the end cell element
                    workSheetPartWriter.WriteEndElement();
                }
                columnNum++;
            }
            else
            {
                throw new InvalidOperationException("There is no active row");
            }
        }


        public void WriteTextRow(IEnumerable<string> texts, int format = 0, bool hidden = false, bool useSharedStrings = false)
        {
            BeginRow(hidden);
            foreach (string text in texts ?? throw new ArgumentNullException(nameof(texts)))
            {
                WriteTextCell(text, format, useSharedStrings);
            }
            EndRow();
        }

        public void WriteNumberRow(IEnumerable<float> numbers, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (float number in numbers ?? throw new ArgumentNullException(nameof(numbers)))
            {
                WriteNumberCell(number, format);
            }
            EndRow();
        }

        public void WriteFormulaRow(IEnumerable<string> formulas, int format = 0, bool hidden = false)
        {
            BeginRow(hidden);
            foreach (string text in formulas ?? throw new ArgumentNullException(nameof(formulas)))
            {
                WriteFormulaCell(text, format);
            }
            EndRow();
        }


        public void AddAutofilter(string range, bool overwrite = false)
        {
            AddAutofilter(new CellRange(range), overwrite);
        }

        public void AddAutofilter(CellRange range, bool overwrite = false)
        {
            if (range == null) { throw new ArgumentNullException(nameof(range)); }
            if ((!overwrite) && SheetAutofilter != null) { throw new InvalidOperationException("There is already a filter in use. Set owerwrite to true to replace it"); }
            if (range.Height != 1) { throw new ArgumentOutOfRangeException(nameof(range), "Range height must be 1"); }
            SheetAutofilter = new AutoFilter() { Reference = range.RangeStringNoSheetName };
        }


        [Obsolete("\"Please use AddListValidator instead.\"", true)]
        public void AddValidator(string range, string formula)
        {
            AddListValidator(range, formula);
        }

        public void AddListValidator(string range,
                                 string formula,
                                 bool allowBlank = true,
                                 bool showInputMessage = true,
                                 bool showErrorMessage = true)
        {
            AddListValidator(new CellRange(range),
                             formula,
                             allowBlank,
                             showInputMessage,
                             showErrorMessage);
        }

        public void AddListValidator(CellRange range,
                             string formula,
                             bool allowBlank = true,
                             bool showInputMessage = true,
                             bool showErrorMessage = true)
        {
            if (range == null) { throw new ArgumentNullException(nameof(range)); }
            if (sheetOpen)
            {
                sheetDataValidations ??= new DataValidations();
                DataValidation dataValidation = new()
                {
                    Type = DataValidationValues.List,
                    AllowBlank = allowBlank,
                    Operator = DataValidationOperatorValues.Equal,
                    ShowInputMessage = showInputMessage,
                    ShowErrorMessage = showErrorMessage,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range.RangeString },
                };

                Formula1 formula1 = new() { Text = formula };

                dataValidation.Append(new[] { formula1 });
                sheetDataValidations.Append(new[] { dataValidation });
                sheetDataValidations.Count = (sheetDataValidations.Count ?? 0) + 1;
            }
            else
            {
                throw new InvalidOperationException("There is no open sheet");
            }
        }

        public void Comment(string text, string reference, string author = "BigExcelCreator")
        {
            if (string.IsNullOrEmpty(author)) { throw new ArgumentOutOfRangeException(nameof(author)); }
            CellRange cellRange = new(reference);
            if (!cellRange.IsSingleCellRange) { throw new ArgumentOutOfRangeException(nameof(reference), $"{nameof(reference)} must be a single cell range"); }
            if (sheetOpen)
            {
                commentManager ??= new();
                commentManager.Add(new CommentReference()
                {
                    Cell = cellRange.RangeStringNoSheetName,
                    Text = text,
                    Author = author,
                });

            }
            else
            {
                throw new InvalidOperationException("There is no open sheet");
            }
        }

        public void AddConditionalFormattingFormula(string reference, string formula, int format)
        {
            CellRange cellRange = new(reference);
            if (formula.IsNullOrWhiteSpace()) { throw new ArgumentNullException(nameof(formula)); }
            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }

            if (!sheetOpen) { throw new InvalidOperationException("There is no open sheet"); }

            ConditionalFormatting conditionalFormatting = new()
            {
                SequenceOfReferences = new(new List<StringValue> { cellRange.RangeStringNoSheetName }),
            };

            ConditionalFormattingRule conditionalFormattingRule = new()
            {
                Type = ConditionalFormatValues.Expression,
                FormatId = (uint)format,
                Priority = conditionalFormattingList.Count + 1,
            };

            conditionalFormattingRule.Append(new[] { new Formula { Text = formula } });

            conditionalFormatting.Append(new[] { conditionalFormattingRule });

            conditionalFormattingList.Add(conditionalFormatting);
        }

        public void AddConditionalFormattingCellIs(string reference, ConditionalFormattingOperatorValues @operator, string value, int format, string value2 = null)
        {
            CellRange cellRange = new(reference);

            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
            if (value.IsNullOrWhiteSpace()) { throw new ArgumentNullException(nameof(value)); }
            if (!sheetOpen) { throw new InvalidOperationException("There is no open sheet"); }
            if (new[] { ConditionalFormattingOperatorValues.Between, ConditionalFormattingOperatorValues.NotBetween }.Contains(@operator)
                && value2.IsNullOrWhiteSpace())
            {
                throw new ArgumentNullException(nameof(value2));
            }

            ConditionalFormatting conditionalFormatting = new()
            {
                SequenceOfReferences = new(new List<StringValue> { cellRange.RangeStringNoSheetName }),
            };

            ConditionalFormattingRule conditionalFormattingRule = new()
            {
                Type = ConditionalFormatValues.CellIs,
                @Operator = @operator,
                FormatId = (uint)format,
                Priority = conditionalFormattingList.Count + 1,
            };

            conditionalFormattingRule.Append(new[] { new Formula { Text = value } });
            if (!value2.IsNullOrWhiteSpace()) { conditionalFormattingRule.Append(new[] { new Formula { Text = value2 } }); }

            conditionalFormatting.Append(new[] { conditionalFormattingRule });

            conditionalFormattingList.Add(conditionalFormatting);
        }

        public void AddConditionalFormattingDuplicatedValues(string reference, int format)
        {
            CellRange cellRange = new(reference);

            if (format < 0) { throw new ArgumentOutOfRangeException(nameof(format)); }
            if (!sheetOpen) { throw new InvalidOperationException("There is no open sheet"); }

            ConditionalFormatting conditionalFormatting = new()
            {
                SequenceOfReferences = new(new List<StringValue> { cellRange.RangeStringNoSheetName }),
            };

            ConditionalFormattingRule conditionalFormattingRule = new()
            {
                Type = ConditionalFormatValues.DuplicateValues,
                FormatId = (uint)format,
                Priority = conditionalFormattingList.Count + 1,
            };

            conditionalFormatting.Append(new[] { conditionalFormattingRule });

            conditionalFormattingList.Add(conditionalFormatting);
        }

        public void MergeCells(CellRange range)
        {
            if (!sheetOpen) { throw new InvalidOperationException("There is no open sheet"); }

            if (SheetMergedCells.Any(range.RangeOverlaps))
            {
                throw new OverlappingRangesException();
            }
            else
            {
                SheetMergedCells.Add(range);
            }
        }

        public void MergeCells(string range)
        {
            MergeCells(new CellRange(range));
        }

        public void CloseDocument()
        {
            if (open)
            {
                if (rowOpen) { EndRow(); }
                if (sheetOpen) { CloseSheet(); }

                WriteSharedStringsPart();
                WriteSheetsAndClosePart();

#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
                Task.WaitAll(SheetTasks.ToArray());
                SheetTasks.Clear();
                Task.WaitAll(DocumentTasks.ToArray());
                DocumentTasks.Clear();
#endif

                Document.Close();

                if (SavingTo == SavingTo.stream)
                {
                    _ = Stream.Seek(0, SeekOrigin.Begin);
                }
            }
            open = false;
        }

        #region IDisposable
        private bool disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                CloseDocument();
                if (disposing)
                {
                    // called via myClass.Dispose(). 
                    // OK to use any private object references
                    workSheetPartWriter?.Dispose();
                    Document.Dispose();
                }
                // Release unmanaged resources.
                // Set large fields to null.                
                disposed = true;
            }
        }

        public void Dispose() // Implement IDisposable
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~BigExcelWriter() // the finalizer
        {
            Dispose(false);
        }
        #endregion

        #region private methods
        private void WriteFilters()
        {
            if (SheetAutofilter == null) { return; }

            workSheetPartWriter.WriteElement(SheetAutofilter);

            SheetAutofilter = null;
        }



        private void WriteValidations()
        {
            if (sheetDataValidations == null) { return; }

            workSheetPartWriter.WriteStartElement(sheetDataValidations);
            foreach (DataValidation dataValidation in sheetDataValidations.ChildElements.Cast<DataValidation>())
            {
                workSheetPartWriter.WriteStartElement(dataValidation);
                workSheetPartWriter.WriteElement(dataValidation.Formula1);
                workSheetPartWriter.WriteEndElement();
            }
            workSheetPartWriter.WriteEndElement();

            sheetDataValidations = null;
        }

        private void WriteConditionalFormatting()
        {
            if (conditionalFormattingList == null || conditionalFormattingList.Count == 0) { return; }

            foreach (ConditionalFormatting conditionalFormatting in conditionalFormattingList)
            {
                workSheetPartWriter.WriteStartElement(conditionalFormatting);
                foreach (ConditionalFormattingRule conditionalFormattingRule in conditionalFormatting.ChildElements.OfType<ConditionalFormattingRule>())
                {
                    workSheetPartWriter.WriteStartElement(conditionalFormattingRule);
                    foreach (var item in conditionalFormattingRule.ChildElements)
                    {
                        workSheetPartWriter.WriteElement(item);
                    }
                    workSheetPartWriter.WriteEndElement();
                }
                workSheetPartWriter.WriteEndElement();
            }

            conditionalFormattingList.Clear();
        }

        private void WriteMergedCells()
        {
            if (SheetMergedCells == null || SheetMergedCells.Count == 0) { return; }

            workSheetPartWriter.WriteStartElement(new MergeCells());

            foreach (CellRange range in SheetMergedCells)
            {
                workSheetPartWriter.WriteElement(new MergeCell { Reference = range.RangeStringNoSheetName });
            }

            workSheetPartWriter.WriteEndElement();

            SheetMergedCells.Clear();
        }

        private int AddTextToSharedStringsTable(string text)
        {
            int pos = SharedStringsList.IndexOf(text);
            if (pos < 0)
            {
                pos = SharedStringsList.Count;
                SharedStringsList.Add(text);
            }
            return pos;
        }

        private void WriteSharedStringsPart()
        {
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
            Task task = new Task(() =>
            {
#endif
                using OpenXmlWriter SharedStringsWriter = OpenXmlWriter.Create(SharedStringTablePart);
                SharedStringsWriter.WriteStartElement(new SharedStringTable());
                foreach (string item in SharedStringsList)
                {
                    SharedStringsWriter.WriteStartElement(new SharedStringItem());
                    SharedStringsWriter.WriteElement(new Text(item));
                    SharedStringsWriter.WriteEndElement();
                }
                SharedStringsWriter.WriteEndElement();
                SharedStringsWriter.Close();
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
            });
            task.Start();
            DocumentTasks.Add(task);
#endif
        }

        private void WriteSheetsAndClosePart()
        {
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
            Task task = new Task(() =>
            {
#endif
                using OpenXmlWriter workbookPartWriter = OpenXmlWriter.Create(Document.WorkbookPart);
                workbookPartWriter.WriteStartElement(new Workbook());
                workbookPartWriter.WriteStartElement(new Sheets());

                foreach (Sheet sheet in sheets)
                {
                    workbookPartWriter.WriteElement(sheet);
                }

                // End Sheets
                workbookPartWriter.WriteEndElement();
                // End Workbook
                workbookPartWriter.WriteEndElement();
                workbookPartWriter.Close();
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
            });
            task.Start();
            DocumentTasks.Add(task);
#endif
        }
        #endregion
    }

    internal enum SavingTo
    {
        file,
        stream,
    }
}
