// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Extensions;
using System;
using System.Globalization;
using System.Text;

namespace BigExcelCreator.Ranges
{
    public class CellRange : IEquatable<CellRange>, IComparable<CellRange>
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

                RangeStringColAndRowPart(sb);

                return sb.ToString();
            }
        }

        public string RangeStringNoSheetName
        {
            get
            {
                StringBuilder sb = new();

                RangeStringColAndRowPart(sb);

                return sb.ToString();
            }
        }

        private void RangeStringColAndRowPart(StringBuilder sb)
        {
            if (StartingColumnIsFixed) { sb.Append('$'); }
            if (StartingColumn != null) { sb.Append(Helpers.GetColumnName(StartingColumn)); }

            if (StartingRowIsFixed) { sb.Append('$'); }
            if (StartingRow != null) { sb.Append(StartingRow); }

            if (!IsSingleCellRange)
            {
                sb.Append(':');

                if (EndingColumnIsFixed) { sb.Append('$'); }
                if (EndingColumn != null) { sb.Append(Helpers.GetColumnName(EndingColumn)); }

                if (EndingRowIsFixed) { sb.Append('$'); }
                if (EndingRow != null) { sb.Append(EndingRow); }
            }
        }

        public int? StartingRow { get; }

        public int? StartingColumn { get; }

        public int? EndingRow { get; }

        public int? EndingColumn { get; }

        public string Sheetname
        {
            get => sheetname;
            set
            {
                if (!value.IsNullOrWhiteSpace() && value.IndexOfAny(invalidSheetCharacters) >= 0)
                {
                    throw new InvalidRangeException();
                }
                else
                {
                    sheetname = value?.Trim();
                }
            }
        }
        private string sheetname;

        public bool StartingRowIsFixed { get; private set; }
        public bool StartingColumnIsFixed { get; private set; }
        public bool EndingRowIsFixed { get; private set; }
        public bool EndingColumnIsFixed { get; private set; }

        public int Width { get => Math.Abs((EndingColumn ?? 0) - (StartingColumn ?? 0)) + 1; }

        public int Height { get => Math.Abs((EndingRow ?? 0) - (StartingRow ?? 0)) + 1; }

        public bool IsSingleCellRange => StartingRow == EndingRow && StartingColumn == EndingColumn;

        private readonly char[] invalidSheetCharacters = @"\/*[]:?".ToCharArray();

        public CellRange(int? column,
                         int? row,
                         string sheetname)
            : this(column, row, column, row, sheetname)
        { }

        public CellRange(int? column,
                         bool fixedColumn,
                         int? row,
                         bool fixedRow,
                         string sheetname)
            : this(column, fixedColumn, row, fixedRow, column, fixedColumn, row, fixedRow, sheetname)
        { }

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
            range = PrepareRangeString(range);

            string possibleRangeValue = SetSheetNameAndGetProbableRange(range);

            string[] rangearray = SplitRangeComponents(possibleRangeValue, out int RANGE_START, out int RANGE_END);

            CountLettersAndNumbers(rangearray[RANGE_START], rangearray[RANGE_END], out int letters1, out int numbers1, out int letters2, out int numbers2);

            rangearray[RANGE_START] = rangearray[RANGE_START].Replace("$", "");
            rangearray[RANGE_END] = rangearray[RANGE_END].Replace("$", "");

            AssertCompleteRange(letters1, numbers1, StartingRowIsFixed, StartingColumnIsFixed, rangearray[RANGE_START]);
            AssertCompleteRange(letters2, numbers2, EndingRowIsFixed, EndingColumnIsFixed, rangearray[RANGE_END]);

            RangeTypes startRangeType = SetRangeType(letters1, numbers1);
            RangeTypes endRangeType = SetRangeType(letters2, numbers2);

            AssertSameRangeType(startRangeType, endRangeType);

            if ((startRangeType & RangeTypes.RowInfinite) == 0)
            {
                StartingRow = int.Parse(rangearray[RANGE_START].Substring(letters1), CultureInfo.InvariantCulture);
            }
            if ((endRangeType & RangeTypes.RowInfinite) == 0)
            {
                EndingRow = int.Parse(rangearray[RANGE_END].Substring(letters2), CultureInfo.InvariantCulture);
            }

            if ((startRangeType & RangeTypes.ColInfinite) == 0)
            {
                StartingColumn = Helpers.GetColumnIndex(rangearray[RANGE_START].Substring(0, rangearray[RANGE_START].Length - numbers1));
            }
            if ((endRangeType & RangeTypes.ColInfinite) == 0)
            {
                EndingColumn = Helpers.GetColumnIndex(rangearray[RANGE_END].Substring(0, rangearray[RANGE_END].Length - numbers2));
            }
        }


        // override object.Equals
        public override bool Equals(object obj)
        {
            return obj is CellRange other && Equals(other);
        }

        public virtual bool Equals(CellRange other)
        {
            return other != null
                && RangeString == other.RangeString
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

        public int CompareTo(CellRange other)
        {
            if (other == null) return 1;
            if (StartingRow < other.StartingRow) return -1;
            if (StartingRow > other.StartingRow) return 1;
            if (StartingColumn < other.StartingColumn) return -1;
            if (StartingColumn > other.StartingColumn) return 1;

            if (EndingRow < other.EndingRow) return -1;
            if (EndingRow > other.EndingRow) return 1;
            if (EndingColumn < other.EndingColumn) return -1;
            if (EndingColumn > other.EndingColumn) return 1;

            return 0;
        }

        public static bool operator ==(CellRange left, CellRange right)
        {
            if (ReferenceEquals(left, null))
            {
                return ReferenceEquals(right, null);
            }

            return left.Equals(right);
        }

        public static bool operator !=(CellRange left, CellRange right)
        {
            return !(left == right);
        }

        public static bool operator <(CellRange left, CellRange right)
        {
            return ReferenceEquals(left, null) ? !ReferenceEquals(right, null) : left.CompareTo(right) < 0;
        }

        public static bool operator <=(CellRange left, CellRange right)
        {
            return ReferenceEquals(left, null) || left.CompareTo(right) <= 0;
        }

        public static bool operator >(CellRange left, CellRange right)
        {
            return !ReferenceEquals(left, null) && left.CompareTo(right) > 0;
        }

        public static bool operator >=(CellRange left, CellRange right)
        {
            return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.CompareTo(right) >= 0;
        }



        private static string PrepareRangeString(string range)
        {
            if (range.IsNullOrWhiteSpace())
            {
                throw new ArgumentNullException(nameof(range));
            }
            else
            {
                return range.Trim();
            }
        }

        private string SetSheetNameAndGetProbableRange(string range)
        {
            string[] firstSplit = range.Split('!');
            switch (firstSplit.Length)
            {
                case 2:
                    if (firstSplit[0].IsNullOrWhiteSpace()) { throw new InvalidRangeException(); }
                    Sheetname = firstSplit[0];
                    return firstSplit[1].ToUpperInvariant();
                case 1:
                    Sheetname = null;
                    return firstSplit[0].ToUpperInvariant();
                default:
                    throw new InvalidRangeException();
            }
        }

        private static string[] SplitRangeComponents(string possibleRangeValue, out int RANGE_START, out int RANGE_END)
        {
            string[] rangearray = possibleRangeValue.Split(':');
            switch (rangearray.Length)
            {
                case 2:
                    RANGE_START = 0;
                    RANGE_END = 1;
                    break;
                case 1:
                    RANGE_START = 0;
                    RANGE_END = 0;
                    break;
                default:
                    throw new InvalidRangeException();
            }
            if (rangearray[RANGE_START].Length == 0 || rangearray[RANGE_END].Length == 0) { throw new InvalidRangeException(); }
            return rangearray;
        }

        private static void AssertCompleteRange(int lettersCount, int NumbersCount, bool rowFixed, bool colFixed, string rangeComponent)
        {
            if (lettersCount + NumbersCount + (rowFixed ? 1 : 0) + (colFixed ? 1 : 0) < rangeComponent.Length)
            { throw new InvalidRangeException(); }
        }

        private static RangeTypes SetRangeType(int lettersCount, int numbersCount)
        {
            RangeTypes rangeType = 0;
            if (lettersCount == 0) { rangeType |= RangeTypes.ColInfinite; }
            if (numbersCount == 0) { rangeType |= RangeTypes.RowInfinite; }

            if (rangeType == (RangeTypes.ColInfinite | RangeTypes.RowInfinite)) { throw new InvalidRangeException(); }

            return rangeType;
        }

        private static void AssertSameRangeType(RangeTypes startRangeType, RangeTypes endRangeType)
        {
            if (startRangeType != endRangeType) { throw new InvalidRangeException(); }
        }

        private void CountLettersAndNumbers(string rangeStart, string rangeEnd, out int letters1, out int numbers1, out int letters2, out int numbers2)
        {
            letters1 = 0;
            letters2 = 0;
            numbers1 = 0;
            numbers2 = 0;
            int i = 0, j = 0;

            if (rangeStart[i] == '$') { StartingColumnIsFixed = true; i++; }
            if (rangeEnd[j] == '$') { EndingColumnIsFixed = true; j++; }
            while (i < rangeStart.Length && char.IsLetter(rangeStart[i])) { letters1++; i++; }
            while (j < rangeEnd.Length && char.IsLetter(rangeEnd[j])) { letters2++; j++; }

            if (i < rangeStart.Length && rangeStart[i] == '$') { StartingRowIsFixed = true; i++; }
            if (j < rangeEnd.Length && rangeEnd[j] == '$') { EndingRowIsFixed = true; j++; }
            while (i < rangeStart.Length && char.IsDigit(rangeStart[i])) { numbers1++; i++; }
            while (j < rangeEnd.Length && char.IsDigit(rangeEnd[j])) { numbers2++; j++; }
        }
    }

    [Flags]
    enum RangeTypes
    {
        None = 0b00,
        ColFinite = None,
        RowFinite = None,
        ColInfinite = 0b01,
        RowInfinite = 0b10,
    }
}
