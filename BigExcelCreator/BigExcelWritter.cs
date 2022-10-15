// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.CommentsManager;
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
using System.Threading;
#if NET40_OR_GREATER || NETSTANDARD1_3_OR_GREATER
using System.Threading.Tasks;
#endif


namespace BigExcelCreator
{
    /// <summary>
    /// This class writes Excel files directly using OpenXML SAX.
    /// Useful when trying to write tens of thousands of rows.
    /// <see cref="https://www.nuget.org/packages/BigExcelCreator/#readme-body-tab">NuGet</see>
    /// <seealso cref="https://github.com/fenase/BigExcelCreator">Source</seealso>
    /// </summary>
    [Obsolete("Use BigExcelwriter instead")]
    public class BigExcelwritter : BigExcelwriter
    {
        public BigExcelwritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType) : base(stream, spreadsheetDocumentType)
        {
        }

        public BigExcelwritter(string path, SpreadsheetDocumentType spreadsheetDocumentType) : base(path, spreadsheetDocumentType)
        {
        }

        public BigExcelwritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet) : base(stream, spreadsheetDocumentType, stylesheet)
        {
        }

        public BigExcelwritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty) : base(stream, spreadsheetDocumentType, skipCellWhenEmpty)
        {
        }

        public BigExcelwritter(string path, SpreadsheetDocumentType spreadsheetDocumentType, Stylesheet stylesheet) : base(path, spreadsheetDocumentType, stylesheet)
        {
        }

        public BigExcelwritter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty) : base(path, spreadsheetDocumentType, skipCellWhenEmpty)
        {
        }

        public BigExcelwritter(string path, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet) : base(path, spreadsheetDocumentType, skipCellWhenEmpty, stylesheet)
        {
        }

        public BigExcelwritter(Stream stream, SpreadsheetDocumentType spreadsheetDocumentType, bool skipCellWhenEmpty, Stylesheet stylesheet) : base(stream, spreadsheetDocumentType, skipCellWhenEmpty, stylesheet)
        {
        }
    }

}
