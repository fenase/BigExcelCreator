// Copyright (c) Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

using BigExcelCreator.Extensions;
using System;
using System.Globalization;
using System.Text;

namespace BigExcelCreator.Ranges
{
    /// <summary>
    /// Range in Excel spreadsheets
    /// </summary>
    public class CellRange : IEquatable<CellRange>, IComparable<CellRange>
    {
        /// <summary>
        /// A <see cref="string"/> representing a range with the sheet's name (if available)
        /// </summary>
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

        /// <summary>
        /// A <see cref="string"/> representing a range without the sheet's name
        /// </summary>
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

            if (!IsSingleCellRange || IsInfiniteCellRange)
            {
                sb.Append(':');

                if (EndingColumnIsFixed) { sb.Append('$'); }
                if (EndingColumn != null) { sb.Append(Helpers.GetColumnName(EndingColumn)); }

                if (EndingRowIsFixed) { sb.Append('$'); }
                if (EndingRow != null) { sb.Append(EndingRow); }
            }
        }

        /// <summary>
        /// Index of the range's first row
        /// </summary>
        public int? StartingRow { get; }

        /// <summary>
        /// Index of the range's first column
        /// </summary>
        public int? StartingColumn { get; }

        /// <summary>
        /// Index of the range's last row
        /// </summary>
        public int? EndingRow { get; }

        /// <summary>
        /// Index of the range's last column
        /// </summary>
        public int? EndingColumn { get; }

        /// <summary>
        /// The name of the range's sheet (if available)
        /// </summary>
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

        /// <summary>
        /// <see langword="true"/> if the starting row is fixed
        /// <para>Represented by '$' in the string representation</para>
        /// </summary>
        public bool StartingRowIsFixed { get; private set; }

        /// <summary>
        /// <see langword="true"/> if the starting column is fixed
        /// <para>Represented by '$' in the string representation</para>
        /// </summary>
        public bool StartingColumnIsFixed { get; private set; }

        /// <summary>
        /// <see langword="true"/> if the ending row is fixed
        /// <para>Represented by '$' in the string representation</para>
        /// </summary>
        public bool EndingRowIsFixed { get; private set; }

        /// <summary>
        /// <see langword="true"/> if the ending column is fixed
        /// <para>Represented by '$' in the string representation</para>
        /// </summary>
        public bool EndingColumnIsFixed { get; private set; }

        /// <summary>
        /// Range's width
        /// </summary>
        public int Width { get => Math.Abs((EndingColumn ?? 0) - (StartingColumn ?? 0)) + 1; }

        /// <summary>
        /// Range's height
        /// </summary>
        public int Height { get => Math.Abs((EndingRow ?? 0) - (StartingRow ?? 0)) + 1; }

        /// <summary>
        /// <see langword="true"/> if the range represents a single column
        /// </summary>
        public bool IsSingleCellRange => StartingRow == EndingRow && StartingColumn == EndingColumn;

        /// <summary>
        /// <see langword="true"/> if the range is infinite in any direction
        /// </summary>
        public bool IsInfiniteCellRange => ((StartRangeType & (RangeTypes.AnyInfinite)) != 0) || ((StartRangeType & RangeTypes.AnyInfinite) != 0);

        /// <summary>
        /// <see langword="true"/> if the range is infinite in any column
        /// </summary>
        public bool IsInfiniteCellRangeCol => ((StartRangeType & (RangeTypes.ColInfinite)) != 0) || ((StartRangeType & RangeTypes.ColInfinite) != 0);

        /// <summary>
        /// <see langword="true"/> if the range is infinite in any row
        /// </summary>
        public bool IsInfiniteCellRangeRow => ((StartRangeType & (RangeTypes.RowInfinite)) != 0) || ((StartRangeType & RangeTypes.RowInfinite) != 0);

        private readonly RangeTypes StartRangeType;

        private readonly char[] invalidSheetCharacters = @"\/*[]:?".ToCharArray();

        /// <summary>
        /// Creates a single cell range using coordinates indexes
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <param name="sheetname"></param>
        /// <exception cref="ArgumentOutOfRangeException">If any index is less than 1</exception>
        /// <exception cref="InvalidRangeException">If a range makes no sense</exception>
        public CellRange(int? column,
                         int? row,
                         string sheetname)
            : this(column, row, column, row, sheetname)
        { }

        /// <summary>
        /// Creates a fixed single cell range using coordinates indexes
        /// </summary>
        /// <param name="column"></param>
        /// <param name="fixedColumn"></param>
        /// <param name="row"></param>
        /// <param name="fixedRow"></param>
        /// <param name="sheetname"></param>
        /// <exception cref="ArgumentOutOfRangeException">If any index is less than 1</exception>
        /// <exception cref="InvalidRangeException">If a range makes no sense</exception>
        public CellRange(int? column,
                         bool fixedColumn,
                         int? row,
                         bool fixedRow,
                         string sheetname)
            : this(column, fixedColumn, row, fixedRow, column, fixedColumn, row, fixedRow, sheetname)
        { }

        /// <summary>
        /// Creates a range using coordinates indexes
        /// </summary>
        /// <param name="startingColumn"></param>
        /// <param name="startingRow"></param>
        /// <param name="endingColumn"></param>
        /// <param name="endingRow"></param>
        /// <param name="sheetname"></param>
        /// <exception cref="ArgumentOutOfRangeException">If any index is less than 1</exception>
        /// <exception cref="InvalidRangeException">If a range makes no sense</exception>
        public CellRange(int? startingColumn,
                         int? startingRow,
                         int? endingColumn,
                         int? endingRow,
                         string sheetname)
            : this(startingColumn, false, startingRow, false, endingColumn, false, endingRow, false, sheetname)
        { }

        /// <summary>
        /// Creates a range using coordinates indexes
        /// </summary>
        /// <param name="startingColumn"></param>
        /// <param name="fixedStartingColumn"></param>
        /// <param name="startingRow"></param>
        /// <param name="fixedStartingRow"></param>
        /// <param name="endingColumn"></param>
        /// <param name="fixedEndingColumn"></param>
        /// <param name="endingRow"></param>
        /// <param name="fixedEndingRow"></param>
        /// <param name="sheetname"></param>
        /// <exception cref="ArgumentOutOfRangeException">If any index is less than 1</exception>
        /// <exception cref="InvalidRangeException">If a range makes no sense</exception>
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

        /// <summary>
        /// Parses a <see cref="string"/> into a range
        /// </summary>
        /// <param name="range"></param>
        /// <exception cref="ArgumentNullException">If <paramref name="range"/> is null</exception>
        /// <exception cref="ArgumentOutOfRangeException">If any index is less than 1</exception>
        /// <exception cref="InvalidRangeException">If a range makes no sense</exception>
        public CellRange(string range)
        {
            range = PrepareRangeString(range);

            string possibleRangeValue = SetSheetNameAndGetProbableRange(range);

            string[] rangearray = SplitRangeComponents(possibleRangeValue, out int RANGE_START, out int RANGE_END);

            CountLettersAndNumbers(rangearray[RANGE_START], rangearray[RANGE_END], out int letters1, out int numbers1, out int letters2, out int numbers2);

#if NET6_0_OR_GREATER
            rangearray[RANGE_START] = rangearray[RANGE_START].Replace("$", "", StringComparison.Ordinal);
            rangearray[RANGE_END] = rangearray[RANGE_END].Replace("$", "", StringComparison.Ordinal);
#else
            rangearray[RANGE_START] = rangearray[RANGE_START].Replace("$", "");
            rangearray[RANGE_END] = rangearray[RANGE_END].Replace("$", "");
#endif

            AssertCompleteRange(letters1, numbers1, StartingRowIsFixed, StartingColumnIsFixed, rangearray[RANGE_START]);
            AssertCompleteRange(letters2, numbers2, EndingRowIsFixed, EndingColumnIsFixed, rangearray[RANGE_END]);

            StartRangeType = SetRangeType(letters1, numbers1);
            RangeTypes EndRangeType = SetRangeType(letters2, numbers2);

            AssertSameRangeType(StartRangeType, EndRangeType);

            if ((StartRangeType & RangeTypes.ColInfinite) == 0)
            {
                StartingRow = int.Parse(rangearray[RANGE_START].Substring(letters1), CultureInfo.InvariantCulture);
            }
            if ((EndRangeType & RangeTypes.ColInfinite) == 0)
            {
                EndingRow = int.Parse(rangearray[RANGE_END].Substring(letters2), CultureInfo.InvariantCulture);
            }

            if ((StartRangeType & RangeTypes.RowInfinite) == 0)
            {
                StartingColumn = Helpers.GetColumnIndex(rangearray[RANGE_START].Substring(0, rangearray[RANGE_START].Length - numbers1));
            }
            if ((EndRangeType & RangeTypes.RowInfinite) == 0)
            {
                EndingColumn = Helpers.GetColumnIndex(rangearray[RANGE_END].Substring(0, rangearray[RANGE_END].Length - numbers2));
            }
        }


        /// <summary>
        /// Range equals
        /// </summary>
        /// <param name="obj">Another range</param>
        /// <returns><see langword="true"/> if ranges are equal. <see langword="false"/> otherwise.</returns>
        public override bool Equals(object obj)
        {
            return obj is CellRange other && Equals(other);
        }

        /// <summary>
        /// Range equals
        /// </summary>
        /// <param name="other">Another range</param>
        /// <returns><see langword="true"/> if ranges are equal. <see langword="false"/> otherwise.</returns>
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

        /// <summary>
        /// Returns the hash code for this range
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hc = 3;
#if NET6_0_OR_GREATER
                hc += 5 * RangeString.GetHashCode(StringComparison.Ordinal);
#else
                hc += 5 * RangeString.GetHashCode();
#endif
                hc += 7 * StartingRow.GetHashCode();
                hc += 5 * EndingRow.GetHashCode();
                hc += 11 * StartingColumn.GetHashCode();
                hc += 13 * EndingColumn.GetHashCode();
                hc += 19 * StartingColumnIsFixed.GetHashCode();
                hc += 23 * EndingColumnIsFixed.GetHashCode();
                hc += 29 * StartingRowIsFixed.GetHashCode();
                hc += 31 * EndingRowIsFixed.GetHashCode();
#if NET6_0_OR_GREATER
                hc += 17 * (Sheetname?.GetHashCode(StringComparison.Ordinal) ?? 0);
#else
                hc += 17 * (Sheetname?.GetHashCode() ?? 0);
#endif
                return hc;
            }
        }

        /// <summary>
        /// Comparison method
        /// </summary>
        /// <param name="other">Another range</param>
        /// <returns>See <see cref="IComparable.CompareTo(object)"/></returns>
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

        /// <summary>
        /// The equality operator
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static bool operator ==(CellRange left, CellRange right)
        {
            if (ReferenceEquals(left, null))
            {
                return ReferenceEquals(right, null);
            }

            return left.Equals(right);
        }

        /// <summary>
        /// The inequality operator
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static bool operator !=(CellRange left, CellRange right)
        {
            return !(left == right);
        }

        /// <summary>
        /// The less than operator
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static bool operator <(CellRange left, CellRange right)
        {
            return ReferenceEquals(left, null) ? !ReferenceEquals(right, null) : left.CompareTo(right) < 0;
        }

        /// <summary>
        /// The less or equal than operator
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static bool operator <=(CellRange left, CellRange right)
        {
            return ReferenceEquals(left, null) || left.CompareTo(right) <= 0;
        }

        /// <summary>
        /// The greater than operator
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static bool operator >(CellRange left, CellRange right)
        {
            return !ReferenceEquals(left, null) && left.CompareTo(right) > 0;
        }

        /// <summary>
        /// The greater or equal than operator
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <returns></returns>
        public static bool operator >=(CellRange left, CellRange right)
        {
            return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.CompareTo(right) >= 0;
        }

        /// <summary>
        /// Compares ranges and returns <see langword="true"/> if they share any cell
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public bool RangeOverlaps(CellRange other)
        {
            if (other == null) { throw new ArgumentNullException(nameof(other)); }
            if (this == other) { return true; }
            if (ColumnOverlaps(other) && RowOverlaps(other)) { return true; }

            return false;
        }

        private bool ColumnOverlaps(CellRange other)
        {
            bool res = false;

            res |= this.IsInfiniteCellRangeCol && other.IsInfiniteCellRangeRow;
            res |= this.IsInfiniteCellRangeRow && other.IsInfiniteCellRangeCol;

            res |= other.IsInfiniteCellRangeCol && other.StartingColumn <= this.StartingColumn && other.EndingColumn >= this.StartingColumn;
            res |= this.IsInfiniteCellRangeCol && this.StartingColumn <= other.StartingColumn && this.EndingColumn >= other.StartingColumn;

            res |= other.StartingColumn.IsBetweenInclusive(this.StartingColumn, this.EndingColumn);
            res |= other.EndingColumn.IsBetweenInclusive(this.StartingColumn, this.EndingColumn);
            res |= this.StartingColumn.IsBetweenInclusive(other.StartingColumn, other.EndingColumn);
            res |= this.EndingColumn.IsBetweenInclusive(other.StartingColumn, other.EndingColumn);

            return res;
        }

        private bool RowOverlaps(CellRange other)
        {
            bool res = false;

            res |= this.IsInfiniteCellRangeCol && other.IsInfiniteCellRangeRow;
            res |= this.IsInfiniteCellRangeRow && other.IsInfiniteCellRangeCol;

            res |= other.IsInfiniteCellRangeRow && other.StartingRow <= this.StartingRow && other.EndingRow >= this.StartingRow;
            res |= this.IsInfiniteCellRangeRow && this.StartingRow <= other.StartingRow && this.EndingRow >= other.StartingRow;

            res |= other.StartingRow.IsBetweenInclusive(this.StartingRow, this.EndingRow);
            res |= other.EndingRow.IsBetweenInclusive(this.StartingRow, this.EndingRow);
            res |= this.StartingRow.IsBetweenInclusive(other.StartingRow, other.EndingRow);
            res |= this.EndingRow.IsBetweenInclusive(other.StartingRow, other.EndingRow);

            return res;
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
            if (lettersCount == 0) { rangeType |= RangeTypes.RowInfinite; }
            if (numbersCount == 0) { rangeType |= RangeTypes.ColInfinite; }

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

        AnyInfinite = ColInfinite | RowInfinite,
    }
}
