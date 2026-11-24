// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.ClassAttributes.Interfaces;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace BigExcelCreator.Helpers
{
    internal static class ExcelValidatorAttributeHelper
    {
        internal static Dictionary<PropertyInfo, IExcelValidatorAttribute> GetValidatorAttributes(
            Type modelType,
            IEnumerable<PropertyInfo> properties)
        {
            Dictionary<PropertyInfo, IExcelValidatorAttribute> validatorAttributes = [];

            foreach (PropertyInfo property in properties)
            {
                IExcelValidatorAttribute[] attributes =
                [
                    .. Attribute
                        .GetCustomAttributes(property, true)
                        .OfType<IExcelValidatorAttribute>()
                ];

                if (attributes.Length > 1)
                {
                    string conflictingTypes = string.Join(
                        ", ",
                        [
                            .. attributes.Select(
                                attribute =>
                                    attribute.GetType().FullName ??
                                    attribute.GetType().Name)
                        ]);

                    throw new InvalidOperationException(
                        string.Format(
                            CultureInfo.InvariantCulture,
                            "Model type '{0}', property '{1}' has multiple Excel validators: {2}. Only one validator is allowed per generated Excel column.",
                            modelType.FullName ?? modelType.Name,
                            property.Name,
                            conflictingTypes));
                }

                if (attributes.Length == 1)
                {
                    validatorAttributes.Add(property, attributes[0]);
                }
            }

            return validatorAttributes;
        }
    }
}