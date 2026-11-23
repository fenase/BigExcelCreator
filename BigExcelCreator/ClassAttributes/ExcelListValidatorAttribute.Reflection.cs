// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace BigExcelCreator.ClassAttributes
{
    public sealed partial class ExcelListValidatorAttribute
    {
        private const BindingFlags MemberLookupFlags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.FlattenHierarchy;
        private const string DisplayAttributeFullName = "System.ComponentModel.DataAnnotations.DisplayAttribute";

        /// <summary>
        /// Extracts all values of the specified <see cref="Enum"/> type and resolves their display or member names.
        /// </summary>
        /// <param name="enumType">The <see cref="Enum"/> type to inspect.</param>
        /// <param name="useDisplayName">Whether to resolve <c>DisplayAttribute.Name</c> when defined.</param>
        /// <returns>An array of display names or member names for all values in <paramref name="enumType"/>.</returns>
        private static string[] ResolveValuesFromEnum(Type enumType, bool useDisplayName)
        {
            Array enumValues = Enum.GetValues(enumType);
            List<string> results = new List<string>(enumValues.Length);
            foreach (object item in enumValues)
            {
                Enum enumItem = (Enum)item;
                string resolvedName = useDisplayName ? GetEnumDisplayName(enumItem) : enumItem.ToString();
                results.Add(resolvedName);
            }

            return results.ToArray();
        }

        /// <summary>
        /// Resolves and validates the collection of string values from the specified static member.
        /// </summary>
        /// <param name="sourceType">The type declaring the static member.</param>
        /// <param name="memberName">The name of the static member.</param>
        /// <param name="useDisplayName">Whether to resolve <c>DisplayAttribute.Name</c> when the member returns <see cref="Enum"/> values.</param>
        /// <returns>An array of strings extracted from the static member.</returns>
        private static string[] ResolveValuesFromMember(Type sourceType, string memberName, bool useDisplayName)
        {
            object rawValue;
            if (!TryGetMemberValue(sourceType, memberName, out rawValue))
            {
                throw new ArgumentException($"Static property, field, or parameterless method '{memberName}' was not found on '{sourceType.FullName}'.", nameof(memberName));
            }

            IEnumerable<string> stringValues = rawValue as IEnumerable<string>;
            if (stringValues != null)
            {
                return stringValues.ToArray();
            }

            IEnumerable enumerable = rawValue as IEnumerable;
            if (enumerable != null && !(rawValue is string))
            {
                return ConvertEnumCollectionToStrings(enumerable, sourceType, memberName, useDisplayName);
            }

            throw new InvalidOperationException($"Member '{memberName}' on '{sourceType.FullName}' must return a non-null IEnumerable<string> or an enum collection.");
        }

        /// <summary>
        /// Attempts to locate and evaluate a static property, field, or parameterless method on the target type.
        /// </summary>
        /// <param name="sourceType">The class or struct type where the static member is declared.</param>
        /// <param name="memberName">The exact name of the static property, field, or method to look up.</param>
        /// <param name="value">
        /// When this method returns <c>true</c>, contains the value returned by the static member (which may be <c>null</c>).
        /// When this method returns <c>false</c>, this is always <c>null</c>.
        /// </param>
        /// <returns><c>true</c> if a matching static member was found and evaluated; otherwise, <c>false</c>.</returns>
        private static bool TryGetMemberValue(Type sourceType, string memberName, out object value)
        {
            PropertyInfo property = sourceType.GetProperty(memberName, MemberLookupFlags);
            if (property != null)
            {
                value = property.GetValue(null, null);
                return true;
            }

            FieldInfo field = sourceType.GetField(memberName, MemberLookupFlags);
            if (field != null)
            {
                value = field.GetValue(null);
                return true;
            }

            MethodInfo method = sourceType.GetMethod(memberName, MemberLookupFlags, null, Type.EmptyTypes, null);
            if (method != null)
            {
                value = method.Invoke(null, null);
                return true;
            }

            value = null;
            return false;
        }

        /// <summary>
        /// Converts a sequence of enum values into their display names or string representations,
        /// throwing if any element is not an <see cref="Enum"/>.
        /// </summary>
        /// <param name="enumerable">The non-generic enumerable sequence to inspect.</param>
        /// <param name="sourceType">The declaring type, used for error reporting.</param>
        /// <param name="memberName">The member name, used for error reporting.</param>
        /// <param name="useDisplayName">Whether to resolve <c>DisplayAttribute.Name</c> when defined.</param>
        /// <returns>An array of resolved enum display names or enum member names.</returns>
        private static string[] ConvertEnumCollectionToStrings(IEnumerable enumerable, Type sourceType, string memberName, bool useDisplayName)
        {
            List<string> results = new List<string>();
            foreach (object item in enumerable)
            {
                Enum enumItem = item as Enum;
                if (enumItem == null)
                {
                    throw new InvalidOperationException($"Member '{memberName}' on '{sourceType.FullName}' must contain only strings or Enum values.");
                }

                string resolvedName = useDisplayName ? GetEnumDisplayName(enumItem) : enumItem.ToString();
                results.Add(resolvedName);
            }

            return results.ToArray();
        }

        /// <summary>
        /// Resolves the <c>DisplayAttribute.Name</c> of an enum member if defined;
        /// otherwise, falls back to <see cref="Enum.ToString()"/>.
        /// </summary>
        /// <param name="enumItem">The enum value to resolve.</param>
        /// <returns>The localized or configured display name if available; otherwise, the enum member name.</returns>
        private static string GetEnumDisplayName(Enum enumItem)
        {
            string memberName = enumItem.ToString();
            Type enumType = enumItem.GetType();
            FieldInfo field = enumType.GetField(memberName);
            if (field == null)
            {
                return memberName;
            }

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

            return memberName;
        }

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
    }
}
