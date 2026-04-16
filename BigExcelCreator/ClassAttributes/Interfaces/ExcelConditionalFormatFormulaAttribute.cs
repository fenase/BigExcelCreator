// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Enums;

namespace BigExcelCreator.ClassAttributes.Interfaces
{
    internal interface IConditionalFormatAttributes
    {
        int Format { get; }
        string StyleName { get; }
        internal StyleModes StyleMode { get; }
    }
}
