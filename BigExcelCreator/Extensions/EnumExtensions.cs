using System;
using System.Globalization;

namespace BigExcelCreator.Extensions
{
    internal static class EnumExtensions
    {
#if NET35
        internal static bool HasFlag<TEnum>(this TEnum @enum, TEnum flag) where TEnum : Enum
        {
            var enumValue = Convert.ToUInt64(@enum, CultureInfo.InvariantCulture);
            var flagValue = Convert.ToUInt64(flag, CultureInfo.InvariantCulture);

            return (enumValue & flagValue) == flagValue;
        }

        public static bool IsEquivalentTo(this Type @type, Type other) => @type == other;
#endif
    }
}