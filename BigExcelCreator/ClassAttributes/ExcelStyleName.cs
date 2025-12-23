using BigExcelCreator.Enums;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Uses a named style for the decorated property.
    /// </summary>
    /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
    /// <param name="useStyleInHeader">Whether the header row will have this style of the one defined for the class with <see cref="ExcelHeaderStyleNameAttribute"/></param>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelStyleNameAttribute(int format, StylingPriority useStyleInHeader = StylingPriority.Header) : Attribute
    {
        /// <summary>
        /// The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.
        /// </summary>
        public int Format { get; } = format;

        /// <summary>
        /// Whether the header row will have this style of the one defined for the class with <see cref="ExcelHeaderStyleNameAttribute"/>
        /// </summary>
        public StylingPriority UseStyleInHeader { get; } = useStyleInHeader;
    }
}
