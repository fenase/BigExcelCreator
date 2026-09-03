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
        lessThan,

        /// <summary>
        /// Less Than Or Equal.
        /// When the item is serialized out as XML, its value is "lessThanOrEqual".
        /// </summary>
        lessThanOrEqual,

        /// <summary>
        /// Equal.
        /// When the item is serialized out as XML, its value is "equal".
        /// </summary>
        equal,

        /// <summary>
        /// Not Equal.
        /// When the item is serialized out as XML, its value is "notEqual".
        /// </summary>
        notEqual,

        /// <summary>
        /// Greater Than Or Equal.
        /// When the item is serialized out as XML, its value is "greaterThanOrEqual".
        /// </summary>
        greaterThanOrEqual,

        /// <summary>
        /// Greater Than.
        /// When the item is serialized out as XML, its value is "greaterThan".
        /// </summary>
        greaterThan,

        /// <summary>
        /// Between.
        /// When the item is serialized out as XML, its value is "between".
        /// </summary>
        between,

        /// <summary>
        /// Not Between.
        /// When the item is serialized out as XML, its value is "notBetween".
        /// </summary>
        notBetween,

        /// <summary>
        /// Contains.
        /// When the item is serialized out as XML, its value is "containsText".
        /// </summary>
        containsText,

        /// <summary>
        /// Does Not Contain.
        /// When the item is serialized out as XML, its value is "notContains".
        /// </summary>
        notContains,

        /// <summary>
        /// Begins With.
        /// When the item is serialized out as XML, its value is "beginsWith".
        /// </summary>
        beginsWith,

        /// <summary>
        /// Ends With.
        /// When the item is serialized out as XML, its value is "endsWith".
        /// </summary>
        endsWith,
    }
}
