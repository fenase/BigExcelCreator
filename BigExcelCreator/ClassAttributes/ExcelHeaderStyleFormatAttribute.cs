using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Format index to apply to the header row.
    /// </summary>
    /// <param name="format">The format index to apply to the cells in header row. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public sealed class ExcelHeaderStyleFormatAttribute(int format) : Attribute
    {
        /// <summary>
        /// The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.
        /// </summary>
        public int Format { get; } = format;
    }
}
