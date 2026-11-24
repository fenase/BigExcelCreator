using System;

namespace BigExcelCreator.Extensions
{
    internal static class SheetExtensions
    {
        internal static string CreateValidatorListSheetName(this string sheetName, string propertyName)
        {
            const int maximumSheetNameLength = 31;
            const int randomPartLength = 12;
            const int remainingBaseLength = maximumSheetNameLength - randomPartLength;

            string sheetPart = CreatePascalCaseValidatorListSheetNamePart(sheetName);
            string propertyPart = CreatePascalCaseValidatorListSheetNamePart(propertyName);
            const int sheetPartLength = remainingBaseLength / 2;
            const int propertyPartLength = remainingBaseLength - sheetPartLength;

            sheetPart = LimitValidatorListSheetNamePart(sheetPart, sheetPartLength);
            propertyPart = LimitValidatorListSheetNamePart(propertyPart, propertyPartLength);

            string randomPart = Guid.NewGuid().ToString("N")[..randomPartLength].ToUpperInvariant();
            string result = sheetPart + propertyPart + randomPart;

            return result.Length > maximumSheetNameLength
                ? result[..maximumSheetNameLength]
                : result;
        }

        private static string CreatePascalCaseValidatorListSheetNamePart(string value)
        {
            if (value.IsNullOrWhiteSpace())
            {
                return "Validation";
            }

            System.Text.StringBuilder result = new();
            bool startWord = true;

            foreach (char character in value)
            {
                if (!char.IsLetterOrDigit(character))
                {
                    startWord = true;
                    continue;
                }

                result.Append(startWord ? char.ToUpperInvariant(character) : character);
                startWord = false;
            }

            return result.Length == 0 ? "Validation" : result.ToString();
        }

        private static string LimitValidatorListSheetNamePart(string value, int maximumLength)
        {
            if (value.Length <= maximumLength)
            {
                return value;
            }

            return value[..maximumLength];
        }
    }
}
