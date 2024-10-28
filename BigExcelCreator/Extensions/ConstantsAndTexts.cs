// Copyright (c) 2022-2024, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

#if NET8_0_OR_GREATER
using System.Text;
#endif

namespace BigExcelCreator.Extensions
{
    internal static class ConstantsAndTexts
    {
        internal const string ConditionalFormattingMustBeOnSheet = "Conditional formatting must be on a sheet";

        internal const string MusBeGreaterThan0 = "must be greater than 0";

        internal const string NoActiveRow = "There is no active row";

        internal const string NeedToBePlacedOnSSheet = "need to be placed on a sheet";

        internal const string MustBeASingleCellRange = "must be a single cell range";

#if NET8_0_OR_GREATER
        internal static readonly CompositeFormat TwoParameterConcatenation = CompositeFormat.Parse("{0}{1}");
        internal static readonly CompositeFormat TwoWordsConcatenation = CompositeFormat.Parse("{0} {1}");
#else
        internal const string TwoParameterConcatenation = "{0}{1}";
        internal const string TwoWordsConcatenation = "{0} {1}";
#endif
    }
}
