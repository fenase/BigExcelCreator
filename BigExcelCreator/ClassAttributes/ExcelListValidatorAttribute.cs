// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes.Interfaces;
using System;
using System.Collections.Generic;

namespace BigExcelCreator.ClassAttributes
{
    /// <summary>
    /// Applies a literal list data validation to a property column.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Values can be supplied as literal strings, all members of an <see cref="Enum"/> type
    /// (resolving <c>DisplayAttribute.Name</c> when present), or dynamically from a static property, field, or parameterless method.
    /// </para>
    /// <code>
    /// internal enum PromotionEligibility
    /// {
    ///     [Display(Name = "Yes, eligible")]
    ///     Yes,
    ///
    ///     [Display(Name = "No, not eligible")]
    ///     No
    /// }
    ///
    /// internal static class ValidationSources
    /// {
    ///     public static IEnumerable&lt;string&gt; Departments =&gt;
    ///         new List&lt;string&gt; { "Engineering", "Finance", "Engineering" }.Distinct();
    /// }
    ///
    /// public sealed class EmployeeRow
    /// {
    ///     // 1. Literal values: ["Active", "Inactive"]
    ///     [ExcelListValidator("Active", "Inactive")]
    ///     public string Status { get; set; }
    ///
    ///     // 2. Enum values using DisplayAttribute.Name: ["Yes, eligible", "No, not eligible"]
    ///     [ExcelListValidator(typeof(PromotionEligibility))]
    ///     public PromotionEligibility CandidateToPromotion { get; set; }
    ///
    ///     // 3. Static member values: ["Engineering", "Finance"]
    ///     [ExcelListValidator(typeof(ValidationSources), nameof(ValidationSources.Departments))]
    ///     public string Department { get; set; }
    /// }
    /// </code>
    /// </remarks>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    [CLSCompliant(false)]
    public sealed partial class ExcelListValidatorAttribute : Attribute, IExcelValidatorAttribute
    {
        /// <summary>
        /// Gets the literal values used by the list data validation.
        /// </summary>
        /// <value>An array of strings representing the allowed cell values.</value>
        public string[] Values { get; }

        /// <summary>
        /// Gets the <see cref="Enum"/> type used to populate the validation values, if specified.
        /// </summary>
        /// <value>The <see cref="Type"/> of the enum, or <c>null</c> when another constructor overload is used.</value>
        public Type EnumType { get; }

        /// <summary>
        /// Gets the type containing the static member that provides the validation values, if specified.
        /// </summary>
        /// <value>The <see cref="Type"/> declaring the static data source, or <c>null</c> when not applicable.</value>
        public Type SourceType { get; }

        /// <summary>
        /// Gets the name of the static member that provides the validation values, if specified.
        /// </summary>
        /// <value>The name of the static property, field, or method, or <c>null</c> when not applicable.</value>
        public string MemberName { get; }

        /// <summary>
        /// Gets a value indicating whether <c>DisplayAttribute.Name</c> metadata is used when resolving enum values.
        /// </summary>
        /// <value><c>true</c> to resolve <c>DisplayAttribute.Name</c> with fallback to <see cref="Enum.ToString()"/>; <c>false</c> to always use <see cref="Enum.ToString()"/>.</value>
        public bool UseDisplayName { get; }

        /// <summary>
        /// Gets a value indicating whether blank values are allowed.
        /// </summary>
        /// <value><c>true</c> if empty cells pass validation; otherwise, <c>false</c>.</value>
        public bool AllowBlank { get; }

        /// <summary>
        /// Gets a value indicating whether the Excel input message is shown.
        /// </summary>
        /// <value><c>true</c> if the input prompt is displayed when selecting the cell; otherwise, <c>false</c>.</value>
        public bool ShowInputMessage { get; }

        /// <summary>
        /// Gets a value indicating whether the Excel error message is shown.
        /// </summary>
        /// <value><c>true</c> if an alert dialog is displayed on invalid entry; otherwise, <c>false</c>.</value>
        public bool ShowErrorMessage { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelListValidatorAttribute"/> class with literal values.
        /// </summary>
        /// <param name="values">The literal values accepted by the validation.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="values"/> is <c>null</c>.</exception>
        public ExcelListValidatorAttribute(params string[] values)
            : this(true, true, true, values)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelListValidatorAttribute"/> class with literal values and display options.
        /// </summary>
        /// <param name="allowBlank">Whether blank values are allowed.</param>
        /// <param name="showInputMessage">Whether the Excel input message is shown.</param>
        /// <param name="showErrorMessage">Whether the Excel error message is shown.</param>
        /// <param name="values">The literal values accepted by the validation.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="values"/> is <c>null</c>.</exception>
        public ExcelListValidatorAttribute(bool allowBlank, bool showInputMessage, bool showErrorMessage, params string[] values)
        {
            Values = values ?? throw new ArgumentNullException(nameof(values));
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelListValidatorAttribute"/> class using all members of an <see cref="Enum"/> type.
        /// </summary>
        /// <param name="enumType">The <see cref="Enum"/> type whose values populate the list.</param>
        /// <param name="useDisplayName">Whether to resolve <c>DisplayAttribute.Name</c> when defined; defaults to <c>true</c>.</param>
        /// <remarks>
        /// <para>
        /// When <paramref name="useDisplayName"/> is <c>true</c>, each enum member is resolved using <c>DisplayAttribute.Name</c> if defined, falling back to <see cref="Enum.ToString()"/>.
        /// </para>
        /// <code>
        /// // Resolves to ["Yes, eligible", "No, not eligible"]
        /// [ExcelListValidator(typeof(PromotionEligibility))]
        /// public PromotionEligibility CandidateToPromotion { get; set; }
        ///
        /// // Resolves to ["Yes", "No"]
        /// [ExcelListValidator(typeof(PromotionEligibility), false)]
        /// public PromotionEligibility RawCandidateToPromotion { get; set; }
        /// </code>
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="enumType"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">Thrown when <paramref name="enumType"/> is not an <see cref="Enum"/> type.</exception>
        public ExcelListValidatorAttribute(Type enumType, bool useDisplayName = true)
            : this(true, true, true, enumType, useDisplayName)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelListValidatorAttribute"/> class using all members of an <see cref="Enum"/> type and display options.
        /// </summary>
        /// <param name="allowBlank">Whether blank values are allowed.</param>
        /// <param name="showInputMessage">Whether the Excel input message is shown.</param>
        /// <param name="showErrorMessage">Whether the Excel error message is shown.</param>
        /// <param name="enumType">The <see cref="Enum"/> type whose values populate the list.</param>
        /// <param name="useDisplayName">Whether to resolve <c>DisplayAttribute.Name</c> when defined; defaults to <c>true</c>.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="enumType"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">Thrown when <paramref name="enumType"/> is not an <see cref="Enum"/> type.</exception>
        public ExcelListValidatorAttribute(bool allowBlank, bool showInputMessage, bool showErrorMessage, Type enumType, bool useDisplayName = true)
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(enumType);
#else
            if (enumType == null)
            {
                throw new ArgumentNullException(nameof(enumType));
            }
#endif
            if (!enumType.IsEnum)
            {
                throw new ArgumentException($"Type '{enumType.FullName}' must be an Enum.", nameof(enumType));
            }

            EnumType = enumType;
            UseDisplayName = useDisplayName;
            Values = ResolveValuesFromEnum(enumType, useDisplayName);
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelListValidatorAttribute"/> class using values resolved from a static member.
        /// </summary>
        /// <param name="sourceType">The <see cref="Type"/> containing the static property, field, or parameterless method.</param>
        /// <param name="memberName">The name of the static member returning an <see cref="IEnumerable{T}"/> of <see cref="string"/> or <see cref="Enum"/> values.</param>
        /// <param name="useDisplayName">Whether to resolve <c>DisplayAttribute.Name</c> when the member returns <see cref="Enum"/> values; defaults to <c>true</c>.</param>
        /// <remarks>
        /// <para>
        /// Supports <c>public</c> and <c>internal</c> static properties, fields, or parameterless methods returning <see cref="IEnumerable{T}"/> of <see cref="string"/> or <see cref="Enum"/>.
        /// </para>
        /// <code>
        /// [ExcelListValidator(typeof(ValidationSources), nameof(ValidationSources.Departments))]
        /// public string Department { get; set; }
        /// </code>
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="sourceType"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">Thrown when <paramref name="memberName"/> is empty, whitespace, or does not exist on <paramref name="sourceType"/>.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the resolved member returns <c>null</c> or an incompatible type.</exception>
        public ExcelListValidatorAttribute(Type sourceType, string memberName, bool useDisplayName = true)
            : this(true, true, true, sourceType, memberName, useDisplayName)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelListValidatorAttribute"/> class using values resolved from a static member and display options.
        /// </summary>
        /// <param name="allowBlank">Whether blank values are allowed.</param>
        /// <param name="showInputMessage">Whether the Excel input message is shown.</param>
        /// <param name="showErrorMessage">Whether the Excel error message is shown.</param>
        /// <param name="sourceType">The <see cref="Type"/> containing the static property, field, or parameterless method.</param>
        /// <param name="memberName">The name of the static member returning an <see cref="IEnumerable{T}"/> of <see cref="string"/> or <see cref="Enum"/> values.</param>
        /// <param name="useDisplayName">Whether to resolve <c>DisplayAttribute.Name</c> when the member returns <see cref="Enum"/> values; defaults to <c>true</c>.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="sourceType"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">Thrown when <paramref name="memberName"/> is empty, whitespace, or does not exist on <paramref name="sourceType"/>.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the resolved member returns <c>null</c> or an incompatible type.</exception>
        public ExcelListValidatorAttribute(bool allowBlank, bool showInputMessage, bool showErrorMessage, Type sourceType, string memberName, bool useDisplayName = true)
        {
            if (string.IsNullOrEmpty(memberName) || memberName.Trim().Length == 0)
            {
                throw new ArgumentException("Member name cannot be null or whitespace.", nameof(memberName));
            }

            SourceType = sourceType ?? throw new ArgumentNullException(nameof(sourceType));
            MemberName = memberName;
            UseDisplayName = useDisplayName;
            Values = ResolveValuesFromMember(sourceType, memberName, useDisplayName);
            AllowBlank = allowBlank;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
        }
    }
}
