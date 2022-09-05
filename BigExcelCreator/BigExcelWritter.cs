using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;

namespace Investigacion.Service.Helper
{
    /// <summary>
    /// Esta clase escribe archivos Excel de forma directa utilizando OpenXML SAX.
    /// Cuando hay decenas de miles de filas, esta es la forma de exportar los datos.
    /// </summary>
    public class BigExcelWritter : IDisposable
    {
        public Stream Stream { get; private set; }

        public SpreadsheetDocumentType SpreadsheetDocumentType { get; private set; }

        public SpreadsheetDocument Document { get; private set; }

        public bool SkipCellWhenEmpty { get; set; }


        private bool sheetOpen = false;
        private string currentSheetName = "";
        private uint currentSheetId = 1;
        private SheetStateValues currentSheetState = SheetStateValues.Visible;
        private bool open = true;
        private int lastRowWritten = 0;
        private bool rowOpen = false;
        private int columnNum = 1;

        private readonly List<Sheet> sheets = new List<Sheet>();

        private DataValidations sheetDataValidations = null;

        private OpenXmlWriter writer;


        private readonly WorkbookPart workbookPart;
        private WorksheetPart workSheetPart;



        public BigExcelWritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType)
        : this(stream, spreadsheetDocumentType, false) { }
        
        public BigExcelWritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet)
        : this(stream, spreadsheetDocumentType, false, stylesheet) { }

        public BigExcelWritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty)
            : this(stream, spreadsheetDocumentType, skipCellWhenEmpty, new Stylesheet()) { }

        public BigExcelWritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet)
        {
            Stream = stream;
            SpreadsheetDocumentType = spreadsheetDocumentType;
            Document = SpreadsheetDocument.Create(Stream, spreadsheetDocumentType);


            workbookPart = Document.AddWorkbookPart();

            if (workbookPart.WorkbookStylesPart == null)
            {
                WorkbookStylesPart wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();
                // add styles to sheet
                wbsp.Stylesheet = stylesheet;
                wbsp.Stylesheet.Save();
            }

            SkipCellWhenEmpty = skipCellWhenEmpty;
        }


        public void CreateAndOpenSheet(string name, List<Column> columns = null,
                                       SheetStateValues sheetState = SheetStateValues.Visible)
        {
            if (!sheetOpen)
            {
                workSheetPart = Document.WorkbookPart.AddNewPart<WorksheetPart>();
                writer = OpenXmlWriter.Create(workSheetPart);
                currentSheetName = name;
                writer.WriteStartElement(new Worksheet());

                if (columns != null && columns.Count > 0)
                {
                    writer.WriteStartElement(new Columns());
                    int indiceColumna = 1;
                    foreach (Column column in columns)
                    {
                        List<OpenXmlAttribute> atributosColumna = new List<OpenXmlAttribute>
                        {
                            new OpenXmlAttribute("min", null, indiceColumna.ToString()),
                            new OpenXmlAttribute("max", null, indiceColumna.ToString()),
                            new OpenXmlAttribute("width", null, column.Width.ToString()),
                            new OpenXmlAttribute("customWidth", null, column.CustomWidth.ToString())
                        };

                        writer.WriteStartElement(new Column(), atributosColumna);
                        writer.WriteEndElement();
                        ++indiceColumna;
                    }
                    writer.WriteEndElement();
                }

                writer.WriteStartElement(new SheetData());
                sheetOpen = true;
                currentSheetState = sheetState;
            }
            else
            {
                throw new InvalidOperationException();
            }
        }

        public void CloseSheet()
        {
            if (sheetOpen)
            {
                // write the end SheetData element
                writer.WriteEndElement();
                // write validations
                WriteValidations();
                // write the end Worksheet element
                writer.WriteEndElement();
                writer.Close();


                sheets.Add(new Sheet()
                {
                    Name = currentSheetName,
                    SheetId = currentSheetId++,
                    Id = Document.WorkbookPart.GetIdOfPart(workSheetPart),
                    State = currentSheetState,
                });


                currentSheetName = "";
                sheetOpen = false;
                workSheetPart = null;
                lastRowWritten = 0;
            }
            else
            {
                throw new InvalidOperationException();
            }
        }



        public void BeginRow(int rownum)
        {
            if (!rowOpen)
            {
                if (rownum > lastRowWritten)
                {
                    lastRowWritten = rownum;
                    //create a new list of attributes
                    List<OpenXmlAttribute> attributes = new List<OpenXmlAttribute>
                    {
                        // add the row index attribute to the list
                        new OpenXmlAttribute("r", null, lastRowWritten.ToString())
                    };

                    //write the row start element with the row index attribute
                    writer.WriteStartElement(new Row(), attributes);
                    rowOpen = true;
                }
                else
                {
                    throw new InvalidOperationException();
                }
            }
            else
            {
                throw new InvalidOperationException();
            }
        }

        public void BeginRow()
        {
            BeginRow(lastRowWritten + 1);
        }


        public void EndRow()
        {
            if (rowOpen)
            {
                // write the end row element
                writer.WriteEndElement();
                columnNum = 1;
                rowOpen = false;
            }
            else
            {
                throw new InvalidOperationException();
            }
        }


        public void WriteTextCell(string text, int formato = 0)
        {
            if (!(SkipCellWhenEmpty && string.IsNullOrEmpty(text)))
            {

                //reset the list of attributes
                List<OpenXmlAttribute> attributes = new List<OpenXmlAttribute>
                {
                    // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                    new OpenXmlAttribute("t", null, "str"),
                    //add the cell reference attribute
                    new OpenXmlAttribute("r", "", string.Format("{0}{1}", GetColumnName(columnNum), lastRowWritten)),
                    //estilos
                    new OpenXmlAttribute("s", null, (formato).ToString())
                };

                //write the cell start element with the type and reference attributes
                writer.WriteStartElement(new Cell(), attributes);
                //write the cell value
                writer.WriteElement(new CellValue(text));

                // write the end cell element
                writer.WriteEndElement();
            }
            columnNum++;
        }


        public void AddValidator(string rango, string formulaFiltro)
        {
            if (sheetOpen)
            {
                sheetDataValidations = sheetDataValidations ?? new DataValidations();
                DataValidation dataValidation = new DataValidation
                {
                    Type = DataValidationValues.List,
                    AllowBlank = true,
                    Operator = DataValidationOperatorValues.Equal,
                    ShowInputMessage = true,
                    ShowErrorMessage = true,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = rango },
                };

                Formula1 formula = new Formula1 { Text = formulaFiltro };

                dataValidation.Append(formula);
                sheetDataValidations.Append(dataValidation);
                sheetDataValidations.Count = (sheetDataValidations.Count ?? 0) + 1;
            }
            else
            {
                throw new InvalidOperationException();
            }
        }



        public void CloseDocument()
        {
            if (open)
            {
                writer = OpenXmlWriter.Create(Document.WorkbookPart);
                writer.WriteStartElement(new Workbook());
                writer.WriteStartElement(new Sheets());


                foreach (Sheet sheet in sheets)
                {
                    writer.WriteElement(sheet);
                }

                // End Sheets
                writer.WriteEndElement();
                // End Workbook
                writer.WriteEndElement();
                writer.Close();
                Document.Close();
            }
            open = false;
        }



        #region IDisposable
        private bool disposed = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                CloseDocument();
                if (disposing)
                {
                    // called via myClass.Dispose(). 
                    // OK to use any private object references
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

        ~BigExcelWritter() // the finalizer
        {
            Dispose(false);
        }
        #endregion


        private void WriteValidations()
        {
            if (sheetDataValidations == null)
            {
                return;
            }

            writer.WriteStartElement(sheetDataValidations);
            foreach (DataValidation item in sheetDataValidations.ChildElements)
            {
                writer.WriteStartElement(item);
                writer.WriteElement(item.Formula1);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();

            sheetDataValidations = null;
        }



        //A simple helper to get the column name from the column index. This is not well tested!
        private static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = string.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }

            return columnName;
        }

    }
}
