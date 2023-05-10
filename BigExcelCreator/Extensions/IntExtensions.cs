// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

namespace BigExcelCreator.Extensions
{
    internal static class IntExtensions
    {
        internal static bool IsBetween(this int? thing, int? min, int? max)
        {
            if (thing == null || min == null || max == null) { return false; }
            return thing > min && thing < max;
        }

        internal static bool IsBetweenInclusive(this int? thing, int? min, int? max)
        {
            if (thing == null || min == null || max == null) { return false; }
            return thing >= min && thing <= max;
        }
    }
}
