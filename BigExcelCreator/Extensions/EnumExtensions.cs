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
            ConditionalFormattingOperator.lessThan => ConditionalFormattingOperatorValues.LessThan,
            ConditionalFormattingOperator.lessThanOrEqual => ConditionalFormattingOperatorValues.LessThanOrEqual,
            ConditionalFormattingOperator.equal => ConditionalFormattingOperatorValues.Equal,
            ConditionalFormattingOperator.notEqual => ConditionalFormattingOperatorValues.NotEqual,
            ConditionalFormattingOperator.greaterThanOrEqual => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
            ConditionalFormattingOperator.greaterThan => ConditionalFormattingOperatorValues.GreaterThan,
            ConditionalFormattingOperator.between => ConditionalFormattingOperatorValues.Between,
            ConditionalFormattingOperator.notBetween => ConditionalFormattingOperatorValues.NotBetween,
            ConditionalFormattingOperator.containsText => ConditionalFormattingOperatorValues.ContainsText,
            ConditionalFormattingOperator.notContains => ConditionalFormattingOperatorValues.NotContains,
            ConditionalFormattingOperator.beginsWith => ConditionalFormattingOperatorValues.BeginsWith,
            ConditionalFormattingOperator.endsWith => ConditionalFormattingOperatorValues.EndsWith,
            _ => throw new ArgumentOutOfRangeException(nameof(@operator), @operator, null)
        };
    }
}