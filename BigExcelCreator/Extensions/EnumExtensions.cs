// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Enums;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;

namespace BigExcelCreator.Extensions
{
    internal static class EnumExtensions
    {
#if NET35
        /// <summary>Determines whether one or more bit fields are set in the current instance.</summary>
        /// <param name="enum"></param>
        /// <param name="flag">An enumeration value.</param>
        /// <returns><see langword="true"/> if the bit field or bit fields that are set in <paramref name="flag"/> are also set in the current instance; otherwise, <see langword="false"/>.</returns>
        /// <exception cref="ArgumentException"><paramref name="flag"/> is a different type than the current instance.</exception>
        internal static bool HasFlag<TEnum>(this TEnum @enum, TEnum flag) where TEnum : Enum
        {
            var enumValue = Convert.ToUInt64(@enum, CultureInfo.InvariantCulture);
            var flagValue = Convert.ToUInt64(flag, CultureInfo.InvariantCulture);

            return (enumValue & flagValue) == flagValue;
        }

        public static bool IsEquivalentTo(this Type @type, Type other) => @type == other;
#endif

        internal static ConditionalFormattingOperatorValues Value(this ConditionalFormattingOperator @operator) => @operator switch
        {
            ConditionalFormattingOperator.LessThan => ConditionalFormattingOperatorValues.LessThan,
            ConditionalFormattingOperator.LessThanOrEqual => ConditionalFormattingOperatorValues.LessThanOrEqual,
            ConditionalFormattingOperator.Equal => ConditionalFormattingOperatorValues.Equal,
            ConditionalFormattingOperator.NotEqual => ConditionalFormattingOperatorValues.NotEqual,
            ConditionalFormattingOperator.GreaterThanOrEqual => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
            ConditionalFormattingOperator.GreaterThan => ConditionalFormattingOperatorValues.GreaterThan,
            ConditionalFormattingOperator.Between => ConditionalFormattingOperatorValues.Between,
            ConditionalFormattingOperator.NotBetween => ConditionalFormattingOperatorValues.NotBetween,
            ConditionalFormattingOperator.ContainsText => ConditionalFormattingOperatorValues.ContainsText,
            ConditionalFormattingOperator.NotContains => ConditionalFormattingOperatorValues.NotContains,
            ConditionalFormattingOperator.BeginsWith => ConditionalFormattingOperatorValues.BeginsWith,
            ConditionalFormattingOperator.EndsWith => ConditionalFormattingOperatorValues.EndsWith,
            _ => throw new ArgumentOutOfRangeException(nameof(@operator), @operator, null)
        };
    }
}