// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: stylesheet

using DocumentFormat.OpenXml.Spreadsheet;

namespace BigExcelCreator.Styles
{
    /// <summary>
    /// A style to be converted to an entry of a stylesheet.
    /// <para>Used in conditional formatting</para>
    /// </summary>
    public class DifferentialStyleElement
    {
        #region props
        /// <summary>
        /// Given name of a differential style
        /// </summary>
        public string Name { get; internal set; }

        /// <summary>
        /// A <see cref="Font"/> to overwrite when the differential format is applied
        /// </summary>
        public Font Font { get; internal set; }

        /// <summary>
        /// A <see cref="Fill"/> to overwrite when the differential format is applied
        /// </summary>
        public Fill Fill { get; internal set; }

        /// <summary>
        /// A <see cref="Border"/> to overwrite when the differential format is applied
        /// </summary>
        public Border Border { get; internal set; }

        /// <summary>
        /// A <see cref="NumberingFormat"/> to overwrite when the differential format is applied
        /// </summary>
        public NumberingFormat NumberingFormat { get; internal set; }

        /// <summary>
        /// A <see cref="Alignment"/> to overwrite when the differential format is applied
        /// </summary>
        public Alignment Alignment { get; internal set; }
        #endregion

        /// <summary>
        /// A <see cref="DifferentialFormat"/> representing this style
        /// </summary>
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
