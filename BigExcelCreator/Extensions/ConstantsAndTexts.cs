#if NET8_0_OR_GREATER
using System.Text;
#endif

namespace BigExcelCreator.Extensions
{
    internal static class ConstantsAndTexts
    {
        internal const string ConditionalFormattingMustBeOnSheet = "Conditional formatting must be on a sheet";

#if NET8_0_OR_GREATER
        internal static readonly CompositeFormat twoParameterConcatenation = CompositeFormat.Parse("{0}{1}");
#else
        internal const string twoParameterConcatenation = "{0}{1}";
#endif
    }
}
