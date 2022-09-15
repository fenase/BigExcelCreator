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

                if (StartingColumnIsFixed) { sb.Append('$'); }
                if (StartingColumn != null) { sb.Append(Helpers.GetColumnName(StartingColumn)); }

                if (StartingRowIsFixed) { sb.Append('$'); }
                if (StartingRow != null) { sb.Append(StartingRow); }

                sb.Append(':');

                if (EndingColumnIsFixed) { sb.Append('$'); }
                if (EndingColumn != null) { sb.Append(Helpers.GetColumnName(EndingColumn)); }

                if (EndingRowIsFixed) { sb.Append('$'); }
                if (EndingRow != null) { sb.Append(EndingRow); }

                return sb.ToString();
            }
        }

        public int? StartingRow { get; }
        public int? StartingColumn { get; }
        public int? EndingRow { get; }
        public int? EndingColumn { get; }
        public string Sheetname { get; }

        public bool StartingRowIsFixed { get; }
        public bool StartingColumnIsFixed { get; }
        public bool EndingRowIsFixed { get; }
        public bool EndingColumnIsFixed { get; }


        private const int RANGE_START = 0;
        private const int RANGE_END = 1;


        public CellRange(int? startingColumn,
                         int? startingRow,
                         int? endingColumn,
                         int? endingRow,
                         string sheetname)
            : this(startingColumn, false, startingRow, false, endingColumn, false, endingRow, false, sheetname)
        { }

        public CellRange(int? startingColumn,
                         bool fixedStartingColumn,
                         int? startingRow,
                         bool fixedStartingRow,
                         int? endingColumn,
                         bool fixedEndingColumn,
                         int? endingRow,
                         bool fixedEndingRow,
                         string sheetname)
        {
            if (startingColumn < 1) { throw new ArgumentOutOfRangeException(nameof(startingColumn)); }
            if (startingRow < 1) { throw new ArgumentOutOfRangeException(nameof(startingRow)); }
            if (endingColumn < 1) { throw new ArgumentOutOfRangeException(nameof(endingColumn)); }
            if (endingRow < 1) { throw new ArgumentOutOfRangeException(nameof(endingRow)); }

            if (startingColumn == null && startingRow == null) { throw new InvalidRangeException(); }
            if (endingColumn == null && endingRow == null) { throw new InvalidRangeException(); }

            if (startingColumn == null ^ endingColumn == null) { throw new InvalidRangeException(); }
            if (startingRow == null ^ endingRow == null) { throw new InvalidRangeException(); }

            Sheetname = sheetname;
            StartingColumn = startingColumn;
            StartingRow = startingRow;
            EndingColumn = endingColumn;
            EndingRow = endingRow;
            StartingColumnIsFixed = fixedStartingColumn;
            StartingRowIsFixed = fixedStartingRow;
            EndingColumnIsFixed = fixedEndingColumn;
            EndingRowIsFixed = fixedEndingRow;
        }


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

            if (rangearray[RANGE_START][i] == '$') { StartingColumnIsFixed = true; i++; }
            if (rangearray[RANGE_END][j] == '$') { EndingColumnIsFixed = true; j++; }
            while (i < rangearray[RANGE_START].Length && char.IsLetter(rangearray[RANGE_START][i])) { letters1++; i++; }
            while (j < rangearray[RANGE_END].Length && char.IsLetter(rangearray[RANGE_END][j])) { letters2++; j++; }

            if (i < rangearray[RANGE_START].Length && rangearray[RANGE_START][i] == '$') { StartingRowIsFixed = true; i++; }
            if (j < rangearray[RANGE_END].Length && rangearray[RANGE_END][j] == '$') { EndingRowIsFixed = true; j++; }
            while (i < rangearray[RANGE_START].Length && char.IsDigit(rangearray[RANGE_START][i])) { numbers1++; i++; }
            while (j < rangearray[RANGE_END].Length && char.IsDigit(rangearray[RANGE_END][j])) { numbers2++; j++; }

            rangearray[RANGE_START] = rangearray[RANGE_START].Replace("$", "");
            rangearray[RANGE_END] = rangearray[RANGE_END].Replace("$", "");

            if (letters1 + numbers1 + (StartingRowIsFixed ? 1 : 0) + (StartingColumnIsFixed ? 1 : 0) < rangearray[RANGE_START].Length)
            { throw new InvalidRangeException(); }
            if (letters2 + numbers2 + (EndingRowIsFixed ? 1 : 0) + (EndingColumnIsFixed ? 1 : 0) < rangearray[RANGE_END].Length)
            { throw new InvalidRangeException(); }

            RangeType startRangeType = 0;
            RangeType endRangeType = 0;
            if (letters1 == 0) { startRangeType |= RangeType.colInfinite; }
            if (letters2 == 0) { endRangeType |= RangeType.colInfinite; }
            if (numbers1 == 0) { startRangeType |= RangeType.rowInfinite; }
            if (numbers2 == 0) { endRangeType |= RangeType.rowInfinite; }

            if (startRangeType == (RangeType.colInfinite | RangeType.rowInfinite)) { throw new InvalidRangeException(); }
            if (endRangeType == (RangeType.colInfinite | RangeType.rowInfinite)) { throw new InvalidRangeException(); }

            if (startRangeType != endRangeType) { throw new InvalidRangeException(); }

            if ((startRangeType & RangeType.rowInfinite) == 0)
            {
                StartingRow = int.Parse(rangearray[RANGE_START].Substring(letters1), CultureInfo.InvariantCulture);
            }
            if ((endRangeType & RangeType.rowInfinite) == 0)
            {
                EndingRow = int.Parse(rangearray[RANGE_END].Substring(letters2), CultureInfo.InvariantCulture);
            }

            if ((startRangeType & RangeType.colInfinite) == 0)
            {
                StartingColumn = Helpers.GetColumnIndex(rangearray[RANGE_START].Substring(0, rangearray[RANGE_START].Length - numbers1));
            }
            if ((endRangeType & RangeType.colInfinite) == 0)
            {
                EndingColumn = Helpers.GetColumnIndex(rangearray[RANGE_END].Substring(0, rangearray[RANGE_END].Length - numbers2));
            }
        }


        // override object.Equals
        public override bool Equals(object obj)
        {
            if (obj == null || obj.GetType() != typeof(CellRange))
            {
                return false;
            }

            CellRange other = obj as CellRange;

            return RangeString == other.RangeString
                && StartingRow == other.StartingRow
                && EndingRow == other.EndingRow
                && StartingColumn == other.StartingColumn
                && EndingColumn == other.EndingColumn
                && Sheetname == other.Sheetname
                && StartingColumnIsFixed == other.StartingColumnIsFixed
                && EndingColumnIsFixed == other.EndingColumnIsFixed
                && StartingRowIsFixed == other.StartingRowIsFixed
                && EndingRowIsFixed == other.EndingRowIsFixed;
        }

        // override object.GetHashCode
        public override int GetHashCode()
        {
            unchecked
            {
                int hc = 3;
                hc += 5 * RangeString.GetHashCode();
                hc += 7 * StartingRow.GetHashCode();
                hc += 5 * EndingRow.GetHashCode();
                hc += 11 * StartingColumn.GetHashCode();
                hc += 13 * EndingColumn.GetHashCode();
                hc += 19 * StartingColumnIsFixed.GetHashCode();
                hc += 23 * EndingColumnIsFixed.GetHashCode();
                hc += 29 * StartingRowIsFixed.GetHashCode();
                hc += 31 * EndingRowIsFixed.GetHashCode();
                hc += 17 * (Sheetname?.GetHashCode() ?? 0);
                return hc;
            }
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
