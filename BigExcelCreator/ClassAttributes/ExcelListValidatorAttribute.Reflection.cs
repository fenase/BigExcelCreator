// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;

namespace BigExcelCreator.ClassAttributes
{
    public sealed partial class ExcelListValidatorAttribute
    {
        private const BindingFlags MemberLookupFlags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.FlattenHierarchy;

        /// <summary>
        /// Extracts all values of the specified <see cref="Enum"/> type and resolves their display or member names.
        /// </summary>
        /// <param name="enumType">The <see cref="Enum"/> type to inspect.</param>
        /// <param name="useDisplayName">Whether to resolve <c>DisplayAttribute.Name</c> when defined.</param>
        /// <returns>An array of display names or member names for all values in <paramref name="enumType"/>.</returns>
        private static string[] ResolveValuesFromEnum(Type enumType, bool useDisplayName)
        {
            Array enumValues = Enum.GetValues(enumType);
            List<string> results = [with(enumValues.Length)];
            foreach (object item in enumValues)
            {
                Enum enumItem = (Enum)item;
                string resolvedName = useDisplayName ? enumItem.GetEnumDisplayName() : enumItem.ToString();
                results.Add(resolvedName);
            }

            return [.. results];
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
            if (!TryGetMemberValue(sourceType, memberName, out object rawValue))
            {
                throw new ArgumentException($"Static property, field, or parameterless method '{memberName}' was not found on '{sourceType.FullName}'.", nameof(memberName));
            }

            if (rawValue is IEnumerable<string> stringValues)
            {
                return [.. stringValues];
            }

            if (rawValue is IEnumerable enumerable && rawValue is not string)
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
            List<string> results = [];
            foreach (object item in enumerable)
            {
                if (item is not Enum enumItem)
                {
                    throw new InvalidOperationException($"Member '{memberName}' on '{sourceType.FullName}' must contain only strings or Enum values.");
                }

                string resolvedName = useDisplayName ? enumItem.GetEnumDisplayName() : enumItem.ToString();
                results.Add(resolvedName);
            }

            return [.. results];
        }
    }
}
