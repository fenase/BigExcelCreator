// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Spreadsheet;

namespace BigExcelCreator.Styles
{
    public class DifferentialStyleElement
    {
        #region props
        public string Name { get; internal set; }

        public Font Font { get; internal set; }
        public Fill Fill { get; internal set; }
        public Border Border { get; internal set; }
        public NumberingFormat NumberingFormat { get; internal set; }
        public Alignment Alignment { get; internal set; }
        #endregion

        public DifferentialFormat DifferentialFormat => new()
        {
            Font = (Font)(Font?.Clone()),
            Fill = (Fill)Fill?.Clone(),
            Border = (Border)Border?.Clone(),
            NumberingFormat = (NumberingFormat)NumberingFormat?.Clone(),
            Alignment = (Alignment)Alignment?.Clone(),
        };
    }
}
