using BigExcelCreator.Enums;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Format to apply to cells.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelStyleFormatAttribute : Attribute
    {
        /// <summary>
        /// The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.
        /// </summary>
        public int Format { get; }

        /// <summary>
        /// The style name to apply to the cell.
        /// </summary>
        public string StyleName { get; }

        /// <summary>
        /// Whether the header row will have this style of the one defined for the class with <see cref="ExcelHeaderStyleFormatAttribute"/>
        /// </summary>
        public StylingPriority UseStyleInHeader { get; }

        internal StyleModes StyleMode { get; }

        /// <summary>
        /// Format index to apply to cells.
        /// </summary>
        /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
        /// <param name="useStyleInHeader">Whether the header row will have this style of the one defined for the class with <see cref="ExcelHeaderStyleFormatAttribute"/></param>
        public ExcelStyleFormatAttribute(int format, StylingPriority useStyleInHeader = StylingPriority.Header)
        {
            Format = format;
            UseStyleInHeader = useStyleInHeader;
            StyleMode = StyleModes.Index;
        }

        /// <summary>
        /// Format to apply to cells.
        /// </summary>
        /// <param name="styleName">The style name to apply to the cell.</param>
        /// <param name="useStyleInHeader">Whether the header row will have this style of the one defined for the class with <see cref="ExcelHeaderStyleFormatAttribute"/></param>
        public ExcelStyleFormatAttribute(string styleName, StylingPriority useStyleInHeader = StylingPriority.Header)
        {
            StyleName = styleName;
            UseStyleInHeader = useStyleInHeader;
            StyleMode = StyleModes.Name;
        }
    }
}
