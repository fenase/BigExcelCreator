// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

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
