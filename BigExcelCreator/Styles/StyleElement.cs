using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace BigExcelCreator.Styles
{
    public class StyleElement
    {
        public string Name { get; set; }

        public CellFormat Style { get; set; }
    }
}
