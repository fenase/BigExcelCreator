namespace BigExcelCreator.Enums
{
    /// <summary>
    /// Conditional formatting operators for Excel conditional formatting rules.
    /// Wrapper for <see cref="DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingOperatorValues"/> to use on attributes.
    /// </summary>
    public enum ConditionalFormattingOperator
    {
        /// <summary>
        /// Less Than.
        /// When the item is serialized out as XML, its value is "lessThan".
        /// </summary>
        LessThan,

        /// <summary>
        /// Less Than Or Equal.
        /// When the item is serialized out as XML, its value is "lessThanOrEqual".
        /// </summary>
        LessThanOrEqual,

        /// <summary>
        /// Equal.
        /// When the item is serialized out as XML, its value is "equal".
        /// </summary>
        Equal,

        /// <summary>
        /// Not Equal.
        /// When the item is serialized out as XML, its value is "notEqual".
        /// </summary>
        NotEqual,

        /// <summary>
        /// Greater Than Or Equal.
        /// When the item is serialized out as XML, its value is "greaterThanOrEqual".
        /// </summary>
        GreaterThanOrEqual,

        /// <summary>
        /// Greater Than.
        /// When the item is serialized out as XML, its value is "greaterThan".
        /// </summary>
        GreaterThan,

        /// <summary>
        /// Between.
        /// When the item is serialized out as XML, its value is "between".
        /// </summary>
        Between,

        /// <summary>
        /// Not Between.
        /// When the item is serialized out as XML, its value is "notBetween".
        /// </summary>
        NotBetween,

        /// <summary>
        /// Contains.
        /// When the item is serialized out as XML, its value is "containsText".
        /// </summary>
        ContainsText,

        /// <summary>
        /// Does Not Contain.
        /// When the item is serialized out as XML, its value is "notContains".
        /// </summary>
        NotContains,

        /// <summary>
        /// Begins With.
        /// When the item is serialized out as XML, its value is "beginsWith".
        /// </summary>
        BeginsWith,

        /// <summary>
        /// Ends With.
        /// When the item is serialized out as XML, its value is "endsWith".
        /// </summary>
        EndsWith,
    }
}
