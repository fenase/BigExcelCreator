using BigExcelCreator.Enums;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Format to apply to the header row.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public sealed class ExcelHeaderStyleFormatAttribute : Attribute
    {
        /// <summary>
        /// The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.
        /// </summary>
        public int Format { get; }

        public string StyleName { get; }

        internal StyleModes StyleMode { get; }

        /// <summary>
        /// Format index to apply to the header row.
        /// </summary>
        /// <param name="format">The format index to apply to the cells in header row. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        public ExcelHeaderStyleFormatAttribute(int format)
        {
            Format = format;
            StyleMode = StyleModes.Index;
        }

        /// <summary>
        /// Format index to apply to the header row.
        /// </summary>
        /// <param name="styleName">The style name to apply to the cell.</param>
        public ExcelHeaderStyleFormatAttribute(string styleName)
        {
            StyleName = styleName;
            StyleMode = StyleModes.Name;
        }
    }
}
