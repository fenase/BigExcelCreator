using BigExcelCreator.Enums;
using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Uses a named style for the decorated property.
    /// </summary>
    /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
    /// <param name="headerStylingPriority">Whether the header row will have this style of the one defined for the class with <see cref="ExcelHeaderStyleNameAttribute"/></param>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelStyleNameAttribute(int format, StylingPriority headerStylingPriority = StylingPriority.Header) : Attribute
    {
        public int Format { get; } = format;

        public StylingPriority HeaderStylingPriority { get; } = headerStylingPriority;
    }
}
