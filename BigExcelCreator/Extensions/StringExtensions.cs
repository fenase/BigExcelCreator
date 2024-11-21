namespace BigExcelCreator.Extensions
{
    internal static class StringExtensions
    {
        internal static bool IsNullOrWhiteSpace(this string value)
#if NET35
            => value == null || string.IsNullOrEmpty(value.Trim());
#else
            => string.IsNullOrWhiteSpace(value);
#endif
    }
}
