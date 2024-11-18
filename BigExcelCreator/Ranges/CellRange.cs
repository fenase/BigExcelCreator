// Copyright (c) 2022-2025, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

// Ignore Spelling: sheetname

using BigExcelCreator.Extensions;
using System;
using System.Globalization;
using System.Text;

namespace BigExcelCreator.Ranges
{
    /// <summary>
    /// Represents a range of cells in an Excel sheet.
    /// </summary>
    /// <remarks>
    /// This class provides properties and methods to handle cell ranges, including their dimensions, overlap checks, and string representations.
    /// </remarks>
    public class CellRange : IEquatable<CellRange>, IComparable<CellRange>
    {
        /// <summary>
        /// Gets the range string representation of the cell range, including the sheet name if available.
        /// </summary>
        /// <value>The range string representation of the cell range, including the sheet name if available.</value>
        public string RangeString
        {
            get
            {
                StringBuilder sb = new();
                if (!Sheetname.IsNullOrWhiteSpace())
                {
                    _ = sb.Append(Sheetname).Append('!');
                }

                RangeStringColAndRowPart(sb);

                return sb.ToString();
            }
        }

        /// <summary>
        /// Gets the range string representation of the cell range without the sheet name.
        /// </summary>
        /// <value>The range string representation of the cell range without the sheet name.</value>
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
            if (StartingColumnIsFixed) { _ = sb.Append('$'); }
            if (StartingColumn != null) { _ = sb.Append(Helpers.GetColumnName(StartingColumn)); }

            if (StartingRowIsFixed) { _ = sb.Append('$'); }
            if (StartingRow != null) { _ = sb.Append(StartingRow); }

            if (!IsSingleCellRange || IsInfiniteCellRange)
            {
                _ = sb.Append(':');

                if (EndingColumnIsFixed) { _ = sb.Append('$'); }
                if (EndingColumn != null) { _ = sb.Append(Helpers.GetColumnName(EndingColumn)); }

                if (EndingRowIsFixed) { _ = sb.Append('$'); }
                if (EndingRow != null) { _ = sb.Append(EndingRow); }
            }
        }

        /// <summary>
        /// Gets the starting row of the cell range.
        /// </summary>
        /// <value>
        /// The starting row of the cell range, or null if the starting row is not specified.
        /// </value>
        public int? StartingRow { get; }

        /// <summary>
        /// Gets the starting column of the cell range.
        /// </summary>
        /// <value>
        /// The starting column of the cell range, or null if the starting column is not specified.
        /// </value>
        public int? StartingColumn { get; }

        /// <summary>
        /// Gets the ending row of the cell range.
        /// </summary>
        /// <value>
        /// The ending row of the cell range, or null if the ending row is not specified.
        /// </value>
        public int? EndingRow { get; }

        /// <summary>
        /// Gets the ending column of the cell range.
        /// </summary>
        /// <value>
        /// The ending column of the cell range, or null if the ending column is not specified.
        /// </value>
        public int? EndingColumn { get; }

        /// <summary>
        /// Gets or sets the sheet name of the cell range.
        /// </summary>
        /// <value>
        /// The sheet name of the cell range.
        /// </value>
        /// <exception cref="InvalidRangeException">
        /// Thrown when the sheet name contains invalid characters.
        /// </exception>
        public string Sheetname
        {
            get => sheetName;
            set
            {
#if NET8_0_OR_GREATER
                if (!value.IsNullOrWhiteSpace() && value.AsSpan().IndexOfAny(invalidSheetCharacters) >= 0)
#else
                if (!value.IsNullOrWhiteSpace() && value.IndexOfAny(invalidSheetCharacters) >= 0)
#endif
                {
                    throw new InvalidRangeException();
                }
                else
                {
                    sheetName = value?.Trim();
                }
            }
        }
        private string sheetName;

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
        /// Gets the width of the cell range.
        /// </summary>
        public int Width => Math.Abs((EndingColumn ?? 0) - (StartingColumn ?? 0)) + 1;

        /// <summary>
        /// Gets the height of the cell range.
        /// </summary>
        public int Height => Math.Abs((EndingRow ?? 0) - (StartingRow ?? 0)) + 1;

        /// <summary>
        /// Gets a value indicating whether the cell range represents a single cell.
        /// </summary>
        /// <value>
        /// True if the cell range represents a single cell; otherwise, false.
        /// </value>
        public bool IsSingleCellRange => StartingRow == EndingRow && StartingColumn == EndingColumn;

        /// <summary>
        /// Gets a value indicating whether the cell range is infinite.
        /// </summary>
        /// <value>
        /// True if the cell range is infinite; otherwise, false.
        /// </value>
        public bool IsInfiniteCellRange => ((StartRangeType & (RangeTypes.AnyInfinite)) != 0) || ((StartRangeType & RangeTypes.AnyInfinite) != 0);

        /// <summary>
        /// Gets a value indicating whether the cell range represents an entire column.
        /// </summary>
        /// <value>
        /// True if the cell range represents an entire column; otherwise, false.
        /// </value>
        public bool IsInfiniteCellRangeCol => ((StartRangeType & (RangeTypes.ColInfinite)) != 0) || ((StartRangeType & RangeTypes.ColInfinite) != 0);

        /// <summary>
        /// Gets a value indicating whether the cell range represents an entire row.
        /// </summary>
        /// <value>
        /// True if the cell range represents an entire row; otherwise, false.
        /// </value>
        public bool IsInfiniteCellRangeRow => ((StartRangeType & (RangeTypes.RowInfinite)) != 0) || ((StartRangeType & RangeTypes.RowInfinite) != 0);

        private readonly RangeTypes StartRangeType;

#if NET8_0_OR_GREATER
        private static readonly System.Buffers.SearchValues<char> invalidSheetCharacters = System.Buffers.SearchValues.Create(@"\/*[]:?");
#else
        private readonly char[] invalidSheetCharacters = @"\/*[]:?".ToCharArray();
#endif

        /// <summary>
        /// Initializes a new instance of the <see cref="CellRange"/> class using coordinates indexes.
        /// <remarks>This creates a single cell range</remarks>
        /// </summary>
        /// <param name="column">The column of the cell range.</param>
        /// <param name="row">The row of the cell range.</param>
        /// <param name="sheetname">The name of the sheet.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when any of the column or row values are less than 1.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the range is invalid.</exception>
        public CellRange(int? column,
                         int? row,
                         string sheetname)
            : this(column, row, column, row, sheetname)
        { }

        /// <summary>
        /// Initializes a new instance of the <see cref="CellRange"/> class using coordinates indexes.
        /// <remarks>This creates a single cell range</remarks>
        /// </summary>
        /// <param name="column">The column of the cell range.</param>
        /// <param name="fixedColumn">Indicates whether the column is fixed.</param>
        /// <param name="row">The row of the cell range.</param>
        /// <param name="fixedRow">Indicates whether the row is fixed.</param>
        /// <param name="sheetname">The name of the sheet.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when any of the column or row values are less than 1.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the range is invalid.</exception>
        public CellRange(int? column,
                         bool fixedColumn,
                         int? row,
                         bool fixedRow,
                         string sheetname)
            : this(column, fixedColumn, row, fixedRow, column, fixedColumn, row, fixedRow, sheetname)
        { }

        /// <summary>
        /// Initializes a new instance of the <see cref="CellRange"/> class using coordinates indexes.
        /// </summary>
        /// <param name="startingColumn">The starting column of the cell range.</param>
        /// <param name="startingRow">The starting row of the cell range.</param>
        /// <param name="endingColumn">The ending column of the cell range.</param>
        /// <param name="endingRow">The ending row of the cell range.</param>
        /// <param name="sheetname">The name of the sheet.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when any of the column or row values are less than 1.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the range is invalid.</exception>
        public CellRange(int? startingColumn,
                         int? startingRow,
                         int? endingColumn,
                         int? endingRow,
                         string sheetname)
            : this(startingColumn, false, startingRow, false, endingColumn, false, endingRow, false, sheetname)
        { }

        /// <summary>
        /// Initializes a new instance of the <see cref="CellRange"/> class using coordinates indexes.
        /// </summary>
        /// <param name="startingColumn">The starting column of the cell range.</param>
        /// <param name="fixedStartingColumn">Indicates whether the starting column is fixed.</param>
        /// <param name="startingRow">The starting row of the cell range.</param>
        /// <param name="fixedStartingRow">Indicates whether the starting row is fixed.</param>
        /// <param name="endingColumn">The ending column of the cell range.</param>
        /// <param name="fixedEndingColumn">Indicates whether the ending column is fixed.</param>
        /// <param name="endingRow">The ending row of the cell range.</param>
        /// <param name="fixedEndingRow">Indicates whether the ending row is fixed.</param>
        /// <param name="sheetname">The name of the sheet.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when any of the column or row values are less than 1.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the range is invalid.</exception>
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
        /// Initializes a new instance of the <see cref="CellRange"/> class from a string representation of a range.
        /// </summary>
        /// <param name="range">The range string to initialize the cell range.</param>
        /// <exception cref="ArgumentNullException">If <paramref name="range"/> is null</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when any of the column or row values are less than 1.</exception>
        /// <exception cref="InvalidRangeException">Thrown when the <paramref name="range"/> does not represent a valid range.</exception>
        public CellRange(string range)
        {
            range = PrepareRangeString(range);

            string possibleRangeValue = SetSheetNameAndGetProbableRange(range);

            string[] rangeArray = SplitRangeComponents(possibleRangeValue, out int RANGE_START, out int RANGE_END);

            CountLettersAndNumbers(rangeArray[RANGE_START], rangeArray[RANGE_END], out int letters1, out int numbers1, out int letters2, out int numbers2);

#if NET6_0_OR_GREATER
            rangeArray[RANGE_START] = rangeArray[RANGE_START].Replace("$", "", StringComparison.Ordinal);
            rangeArray[RANGE_END] = rangeArray[RANGE_END].Replace("$", "", StringComparison.Ordinal);
#else
            rangeArray[RANGE_START] = rangeArray[RANGE_START].Replace("$", "");
            rangeArray[RANGE_END] = rangeArray[RANGE_END].Replace("$", "");
#endif

            AssertCompleteRange(letters1, numbers1, StartingRowIsFixed, StartingColumnIsFixed, rangeArray[RANGE_START]);
            AssertCompleteRange(letters2, numbers2, EndingRowIsFixed, EndingColumnIsFixed, rangeArray[RANGE_END]);

            StartRangeType = SetRangeType(letters1, numbers1);
            RangeTypes EndRangeType = SetRangeType(letters2, numbers2);

            AssertSameRangeType(StartRangeType, EndRangeType);

            if ((StartRangeType & RangeTypes.ColInfinite) == 0)
            {
#if NET6_0_OR_GREATER
                StartingRow = int.Parse(rangeArray[RANGE_START].AsSpan(letters1), provider: CultureInfo.InvariantCulture);
#else
                StartingRow = int.Parse(rangeArray[RANGE_START].Substring(letters1), CultureInfo.InvariantCulture);
#endif
            }
            if ((EndRangeType & RangeTypes.ColInfinite) == 0)
            {
#if NET6_0_OR_GREATER
                EndingRow = int.Parse(rangeArray[RANGE_END].AsSpan(letters2), provider: CultureInfo.InvariantCulture);
#else
                EndingRow = int.Parse(rangeArray[RANGE_END].Substring(letters2), CultureInfo.InvariantCulture);
#endif
            }

            if ((StartRangeType & RangeTypes.RowInfinite) == 0)
            {
#if NET6_0_OR_GREATER
                StartingColumn = Helpers.GetColumnIndex(rangeArray[RANGE_START][..^numbers1]);
#else
                StartingColumn = Helpers.GetColumnIndex(rangeArray[RANGE_START].Substring(0, rangeArray[RANGE_START].Length - numbers1));
#endif
            }
            if ((EndRangeType & RangeTypes.RowInfinite) == 0)
            {
#if NET6_0_OR_GREATER
                EndingColumn = Helpers.GetColumnIndex(rangeArray[RANGE_END][..^numbers2]);
#else
                EndingColumn = Helpers.GetColumnIndex(rangeArray[RANGE_END].Substring(0, rangeArray[RANGE_END].Length - numbers2));
#endif
            }
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current <see cref="CellRange"/> instance.
        /// </summary>
        /// <param name="obj">The object to compare with the current instance.</param>
        /// <returns>
        /// True if the specified object is a <see cref="CellRange"/> and is equal to the current instance; otherwise, false.
        /// </returns>
        public override bool Equals(object obj) => obj is CellRange other && Equals(other);

        /// <summary>
        /// Determines whether the specified <see cref="CellRange"/> is equal to the current <see cref="CellRange"/> instance.
        /// </summary>
        /// <param name="other">The <see cref="CellRange"/> to compare with the current instance.</param>
        /// <returns>
        /// True if the specified <see cref="CellRange"/> is equal to the current instance; otherwise, false.
        /// </returns>
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
        /// Returns the hash code for this instance.
        /// </summary>
        /// <returns>
        /// A 32-bit signed integer hash code.
        /// </returns>
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
        /// Compares the current instance with another <see cref="CellRange"/> and returns an integer that indicates whether the current instance precedes,
        /// follows, or occurs in the same position in the sort order as the other <see cref="CellRange"/>.
        /// </summary>
        /// <param name="other">The <see cref="CellRange"/> to compare with the current instance.</param>
        /// <returns>
        /// A value that indicates the relative order of the objects being compared. The return value has these meanings:
        /// <list type="bullet">
        /// <item>
        /// <description>Less than zero: This instance precedes <paramref name="other"/> in the sort order.</description>
        /// </item>
        /// <item>
        /// <description>Zero: This instance occurs in the same position in the sort order as <paramref name="other"/>.</description>
        /// </item>
        /// <item>
        /// <description>Greater than zero: This instance follows <paramref name="other"/> in the sort order.</description>
        /// </item>
        /// </list>
        /// </returns>
        public int CompareTo(CellRange other)
        {
            if (other == null) { return 1; }
            if (StartingRow < other.StartingRow) { return -1; }
            if (StartingRow > other.StartingRow) { return 1; }
            if (StartingColumn < other.StartingColumn) { return -1; }
            if (StartingColumn > other.StartingColumn) { return 1; }

            if (EndingRow < other.EndingRow) { return -1; }
            if (EndingRow > other.EndingRow) { return 1; }
            if (EndingColumn < other.EndingColumn) { return -1; }
            if (EndingColumn > other.EndingColumn) { return 1; }

            return 0;
        }

        /// <summary>
        /// Determines whether two specified <see cref="CellRange"/> objects have the same value.
        /// </summary>
        /// <param name="left">The first <see cref="CellRange"/> to compare.</param>
        /// <param name="right">The second <see cref="CellRange"/> to compare.</param>
        /// <returns>
        /// <c>true</c> if the value of <paramref name="left"/> is the same as the value of <paramref name="right"/>; otherwise, <c>false</c>.
        /// </returns>
        public static bool operator ==(CellRange left, CellRange right)
        {
            if (left is null)
            {
                return right is null;
            }

            return left.Equals(right);
        }

        /// <summary>
        /// Determines whether two specified <see cref="CellRange"/> objects have different values.
        /// </summary>
        /// <param name="left">The first <see cref="CellRange"/> to compare.</param>
        /// <param name="right">The second <see cref="CellRange"/> to compare.</param>
        /// <returns>
        /// <c>true</c> if the value of <paramref name="left"/> is different from the value of <paramref name="right"/>; otherwise, <c>false</c>.
        /// </returns>
        public static bool operator !=(CellRange left, CellRange right)
        {
            return !(left == right);
        }

        /// <summary>
        /// Determines whether one specified <see cref="CellRange"/> is less than another specified <see cref="CellRange"/>.
        /// </summary>
        /// <param name="left">The first <see cref="CellRange"/> to compare.</param>
        /// <param name="right">The second <see cref="CellRange"/> to compare.</param>
        /// <returns>
        /// <c>true</c> if the value of <paramref name="left"/> is less than the value of <paramref name="right"/>; otherwise, <c>false</c>.
        /// </returns>
        public static bool operator <(CellRange left, CellRange right)
        {
            return left is null ? right is not null : left.CompareTo(right) < 0;
        }

        /// <summary>
        /// Determines whether one specified <see cref="CellRange"/> is less than or equal to another specified <see cref="CellRange"/>.
        /// </summary>
        /// <param name="left">The first <see cref="CellRange"/> to compare.</param>
        /// <param name="right">The second <see cref="CellRange"/> to compare.</param>
        /// <returns>
        /// <c>true</c> if the value of <paramref name="left"/> is less than or equal to the value of <paramref name="right"/>; otherwise, <c>false</c>.
        /// </returns>
        public static bool operator <=(CellRange left, CellRange right)
        {
            return left is null || left.CompareTo(right) <= 0;
        }

        /// <summary>
        /// Determines whether one specified <see cref="CellRange"/> is greater than another specified <see cref="CellRange"/>.
        /// </summary>
        /// <param name="left">The first <see cref="CellRange"/> to compare.</param>
        /// <param name="right">The second <see cref="CellRange"/> to compare.</param>
        /// <returns>
        /// <c>true</c> if the value of <paramref name="left"/> is greater than the value of <paramref name="right"/>; otherwise, <c>false</c>.
        /// </returns>
        public static bool operator >(CellRange left, CellRange right)
        {
            return left is not null && left.CompareTo(right) > 0;
        }

        /// <summary>
        /// Determines whether one specified <see cref="CellRange"/> is greater than or equal to another specified <see cref="CellRange"/>.
        /// </summary>
        /// <param name="left">The first <see cref="CellRange"/> to compare.</param>
        /// <param name="right">The second <see cref="CellRange"/> to compare.</param>
        /// <returns>
        /// <c>true</c> if the value of <paramref name="left"/> is greater than or equal to the value of <paramref name="right"/>; otherwise, <c>false</c>.
        /// </returns>
        public static bool operator >=(CellRange left, CellRange right)
        {
            return left is null ? right is null : left.CompareTo(right) >= 0;
        }

        /// <summary>
        /// Determines whether the current <see cref="CellRange"/> overlaps with another specified <see cref="CellRange"/>.
        /// </summary>
        /// <param name="other">The <see cref="CellRange"/> to compare with the current <see cref="CellRange"/>.</param>
        /// <returns>
        /// <c>true</c> if the current <see cref="CellRange"/> overlaps with the <paramref name="other"/> <see cref="CellRange"/>; otherwise, <c>false</c>.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="other"/> is <c>null</c>.</exception>
        public bool RangeOverlaps(CellRange other)
        {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(other);
#else
            if (other == null) { throw new ArgumentNullException(nameof(other)); }
#endif
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
            string[] rangeArray = possibleRangeValue.Split(':');
            switch (rangeArray.Length)
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
            if (rangeArray[RANGE_START].Length == 0 || rangeArray[RANGE_END].Length == 0) { throw new InvalidRangeException(); }
            return rangeArray;
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
