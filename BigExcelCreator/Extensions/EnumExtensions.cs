// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

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