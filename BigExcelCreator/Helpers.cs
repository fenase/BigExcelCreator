using BigExcelCreator.Extensions;
using System;
using System.Collections.Generic;
using System.Text;

namespace BigExcelCreator
{
    internal static class Helpers
    {
        //A simple helper to get the column name from the column index. This is not well tested!
        //Starts at 1
        internal static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            StringBuilder columnName = new();
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                _ = columnName.Insert(0, Convert.ToChar(65 + modifier));
                dividend = (dividend - modifier) / 26;
            }

            return columnName.ToString();
        }
        internal static string GetColumnName(int? columnIndex)
            => columnIndex == null ? string.Empty : GetColumnName(columnIndex.Value);

        private static readonly List<char> chars = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

        internal static int GetColumnIndex(string columnName)
        {
            return columnName.IsNullOrWhiteSpace()
                ? 0
                : ((chars.IndexOf(char.ToUpperInvariant(columnName[0])) + 1) * (int)Math.Pow(chars.Count, columnName.Length - 1))
                    + GetColumnIndex(columnName.Substring(1));
        }
    }
}
