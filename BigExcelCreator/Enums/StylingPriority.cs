// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

namespace BigExcelCreator.Enums
{
    /// <summary>
    /// Specifies, for a column, whether the header row or the data rows have priority when applying styles.
    /// </summary>
    public enum StylingPriority
    {
        /// <summary>
        /// Prefer styles applied to the header row (if present) over styles applied to data rows.
        /// </summary>
        Header,

        /// <summary>
        /// Prefer styles applied to data rows over styles applied to the header row.
        /// </summary>
        Data,
    }
}
