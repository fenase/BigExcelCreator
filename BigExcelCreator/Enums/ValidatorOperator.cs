// Copyright (c) 2022-2026, Federico Seckel.
// Licensed under the BSD 3-Clause License. See LICENSE file in the project root for full license information.

namespace BigExcelCreator.Enums
{
    /// <summary>Specifies the comparison operator used by numeric data validations.</summary>
    public enum ValidatorOperator
    {
        /// <summary>Both operands must be equal.</summary>
        Equal,

        /// <summary>The operands must differ.</summary>
        NotEqual,

        /// <summary>The value must be greater than the first operand.</summary>
        GreaterThan,

        /// <summary>The value must be greater than or equal to the first operand.</summary>
        GreaterThanOrEqual,

        /// <summary>The value must be less than the first operand.</summary>
        LessThan,

        /// <summary>The value must be less than or equal to the first operand.</summary>
        LessThanOrEqual,

        /// <summary>The value must be between the first and second operands.</summary>
        Between,

        /// <summary>The value must not be between the first and second operands.</summary>
        NotBetween,
    }
}
