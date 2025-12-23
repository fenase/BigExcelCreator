using BigExcelCreator.Enums;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Format index to apply to cells.
    /// </summary>
    /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
    /// <param name="useStyleInHeader">Whether the header row will have this style of the one defined for the class with <see cref="ExcelHeaderStyleFormatAttribute"/></param>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelStyleFormatAttribute(int format, StylingPriority useStyleInHeader = StylingPriority.Header) : Attribute
    {
        /// <summary>
        /// The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.
        /// </summary>
        public int Format { get; } = format;

        /// <summary>
        /// Whether the header row will have this style of the one defined for the class with <see cref="ExcelHeaderStyleFormatAttribute"/>
        /// </summary>
        public StylingPriority UseStyleInHeader { get; } = useStyleInHeader;
    }
}
