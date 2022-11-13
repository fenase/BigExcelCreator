namespace BigExcelCreator.Extensions
{
    internal static class StringExtensions
    {
        internal static bool IsNullOrWhiteSpace(this string value)
        {
#if NET35
            if (value == null) { return true; }
            return string.IsNullOrEmpty(value.Trim());
#else
            return string.IsNullOrWhiteSpace(value);
#endif
        }
    }
}
