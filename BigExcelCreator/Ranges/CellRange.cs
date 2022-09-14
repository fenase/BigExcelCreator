using BigExcelCreator.Extensions;
using System;
using System.Globalization;
using System.Text;

namespace BigExcelCreator.Ranges
{
    public class CellRange
    {
        public string RangeString
        {
            get
            {
                StringBuilder sb = new();
                if (!Sheetname.IsNullOrWhiteSpace())
                {
                    sb.Append(Sheetname).Append('!');
                }

                if (FixStartCol) { sb.Append('$'); }
                sb.Append(Helpers.GetColumnName(RangeStartCol));

                if (FixStartRow) { sb.Append('$'); }
                sb.Append(RangeStartRow);

                sb.Append(':');

                if (FixEndCol) { sb.Append('$'); }
                sb.Append(Helpers.GetColumnName(RangeEndCol));

                if (FixEndRow) { sb.Append('$'); }
                sb.Append(RangeEndRow);

                return sb.ToString();
            }
        }

        public int RangeStartRow { get; }
        public int RangeStartCol { get; }
        public int RangeEndRow { get; }
        public int RangeEndCol { get; }
        public string Sheetname { get; }

        private bool FixStartRow { get; }
        private bool FixStartCol { get; }
        private bool FixEndRow { get; }
        private bool FixEndCol { get; }


        private const int RANGE_START = 0;
        private const int RANGE_END = 1;


        public CellRange(string range)
        {
            if (range.IsNullOrWhiteSpace())
            {
                throw new ArgumentNullException(nameof(range));
            }
            else
            {
                range = range.Trim();
            }

            string[] firstSplit = range.Split('!');
            string possibleRangeValue;
            switch (firstSplit.Length)
            {
                case 2:
                    if (firstSplit[0].IsNullOrWhiteSpace()) { throw new InvalidRangeException(); }
                    Sheetname = firstSplit[0];
                    possibleRangeValue = firstSplit[1].ToUpperInvariant();
                    break;
                case 1:
                    Sheetname = null;
                    possibleRangeValue = firstSplit[0].ToUpperInvariant();
                    break;
                default:
                    throw new InvalidRangeException();
            }

            string[] rangearray = possibleRangeValue.Split(':');
            if (rangearray.Length != 2) { throw new InvalidRangeException(); }
            if (rangearray[RANGE_START].Length == 0 || rangearray[RANGE_END].Length == 0) { throw new InvalidRangeException(); }

            int letters1 = 0, letters2 = 0, numbers1 = 0, numbers2 = 0;

            int i = 0, j = 0;

            if (rangearray[RANGE_START][i] == '$') { FixStartCol = true; i++; }
            if (rangearray[RANGE_END][j] == '$') { FixEndCol = true; j++; }
            while (i < rangearray[RANGE_START].Length && char.IsLetter(rangearray[RANGE_START][i])) { letters1++; i++; }
            while (j < rangearray[RANGE_END].Length && char.IsLetter(rangearray[RANGE_END][j])) { letters2++; j++; }

            if (i < rangearray[RANGE_START].Length && rangearray[RANGE_START][i] == '$') { FixStartRow = true; i++; }
            if (j < rangearray[RANGE_END].Length && rangearray[RANGE_END][j] == '$') { FixEndRow = true; j++; }
            while (i < rangearray[RANGE_START].Length && char.IsDigit(rangearray[RANGE_START][i])) { numbers1++; i++; }
            while (j < rangearray[RANGE_END].Length && char.IsDigit(rangearray[RANGE_END][j])) { numbers2++; j++; }

            rangearray[RANGE_START] = rangearray[RANGE_START].Replace("$", "");
            rangearray[RANGE_END] = rangearray[RANGE_END].Replace("$", "");

            if (letters1 + numbers1 + (FixStartRow ? 1 : 0) + (FixStartCol ? 1 : 0) < rangearray[RANGE_START].Length)
            { throw new InvalidRangeException(); }
            if (letters2 + numbers2 + (FixEndRow ? 1 : 0) + (FixEndCol ? 1 : 0) < rangearray[RANGE_END].Length)
            { throw new InvalidRangeException(); }

            RangeType startRangeType = 0;
            RangeType endRangeType = 0;
            if (letters1 == 0) { startRangeType |= RangeType.rowInfinite; }
            if (letters2 == 0) { endRangeType |= RangeType.rowInfinite; }
            if (numbers1 == 0) { startRangeType |= RangeType.colInfinite; }
            if (numbers2 == 0) { endRangeType |= RangeType.colInfinite; }

            if (startRangeType == (RangeType.colInfinite | RangeType.rowInfinite)) { throw new InvalidRangeException(); }
            if (endRangeType == (RangeType.colInfinite | RangeType.rowInfinite)) { throw new InvalidRangeException(); }

            if (startRangeType != endRangeType) { throw new InvalidRangeException(); }


            RangeStartRow = int.Parse(rangearray[RANGE_START].Substring(letters1), CultureInfo.InvariantCulture);
            RangeEndRow = int.Parse(rangearray[RANGE_END].Substring(letters2), CultureInfo.InvariantCulture);

            RangeStartCol = Helpers.GetColumnIndex(rangearray[RANGE_START].Substring(0, letters1));
            RangeEndCol = Helpers.GetColumnIndex(rangearray[RANGE_END].Substring(0, letters2));
        }
    }

    [Flags]
    enum RangeType
    {
        colFinite = 0b00,
        rowFinite = colFinite,
        colInfinite = 0b01,
        rowInfinite = 0b10,
    }
}
