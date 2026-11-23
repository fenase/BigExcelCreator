// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Enums;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;
using System.Reflection;

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

        /// <summary>
        /// Resolves the <c>DisplayAttribute.Name</c> of an enum member if defined;
        /// otherwise, falls back to <see cref="Enum.ToString()"/>.
        /// </summary>
        /// <param name="enumItem">The enum value to resolve.</param>
        /// <returns>The localized or configured display name if available; otherwise, the enum member name.</returns>
        internal static string GetEnumDisplayName(this Enum enumItem)
        {
            string memberName = enumItem.ToString();
            Type enumType = enumItem.GetType();
            FieldInfo field = enumType.GetField(memberName);
            if (field == null)
            {
                return memberName;
            }

#if NET40_OR_GREATER || NETCOREAPP
            object[] attributes = field.GetCustomAttributes(typeof(System.ComponentModel.DataAnnotations.DisplayAttribute), false);
            foreach (object attribute in attributes)
            {
                string displayName = ((System.ComponentModel.DataAnnotations.DisplayAttribute)attribute).GetName();
                if (!string.IsNullOrEmpty(displayName))
                {
                    return displayName;
                }
            }
#else
            object[] attributes = field.GetCustomAttributes(false);
            foreach (object attribute in attributes)
            {
                Type attributeType = attribute.GetType();
                if (attributeType.FullName == DisplayAttributeFullName)
                {
                    string displayName = ExtractDisplayAttributeName(attribute, attributeType);
                    if (!string.IsNullOrEmpty(displayName))
                    {
                        return displayName;
                    }
                }
            }
#endif

            return memberName;
        }

#if !(NET40_OR_GREATER || NETCOREAPP)
        private const string DisplayAttributeFullName = "System.ComponentModel.DataAnnotations.DisplayAttribute";

        /// <summary>
        /// Extracts the resolved display name from a <c>DisplayAttribute</c> instance without requiring a compile-time assembly reference.
        /// </summary>
        /// <param name="displayAttribute">The <c>DisplayAttribute</c> instance.</param>
        /// <param name="attributeType">The runtime <see cref="Type"/> of the attribute.</param>
        /// <returns>The resolved display name, or <c>null</c> if not set.</returns>
        private static string ExtractDisplayAttributeName(object displayAttribute, Type attributeType)
        {
            MethodInfo getNameMethod = attributeType.GetMethod("GetName", BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
            if (getNameMethod != null)
            {
                return getNameMethod.Invoke(displayAttribute, null) as string;
            }

            PropertyInfo nameProperty = attributeType.GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
            if (nameProperty != null)
            {
                return nameProperty.GetValue(displayAttribute, null) as string;
            }

            return null;
        }
#endif
    }
}