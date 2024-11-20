using System;

namespace BigExcelCreator.Ranges
{
    [Flags]
    internal enum RangeTypes
    {
        None = 0b00,
        ColFinite = None,
        RowFinite = None,
        ColInfinite = 0b01,
        RowInfinite = 0b10,

        AnyInfinite = ColInfinite | RowInfinite,
    }
}
