using System;

namespace BigExcelCreator.Enums
{
    [Flags]
    internal enum StyleModes
    {
        None = 0,
        Index = 1,
        Name = 2,
        Both = Index | Name,
    }
}
