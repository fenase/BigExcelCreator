using System;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Uses a named style for the header row of the decorated class.
    /// </summary>
    /// <param name="format">The format index to apply to the cell. Default is 0. See <see cref="Styles.StyleList.GetIndexByName(string)"/>.</param>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public sealed class ExcelHeaderStyleNameAttribute(int format) : Attribute
    {
        public int Format { get; } = format;
    }
}
