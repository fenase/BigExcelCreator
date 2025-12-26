# Changelog

## 3.4
THIS IS A BETA VERSION. It may contain bugs and breaking changes.
### Added
- Ability to reference styles by name when writing cells. This is achieved by passing a `StyleList` instance when creating the `BigExcelWriter` instead of a `StyleSheet` instance.

## 3.3 (and 2.3.2025.35601)
### Added
- Support to write a new sheet directly from a list of objects of a custom class, configuring which properties to write using attributes.
- .Net10.0 build target
### Removed
- .NetStandard1.3 build target (from version 2.3.x) since it's EOL.


## 3.2 (and 2.3.2024.32600)
### Added
- Throw exceptions when attempting to create a sheet with an invalid name.
- Throw exceptions when attempting to create a sheet with the same name as an existing sheet.
### Changed
- It's no longer mandatory to provide a SpreadsheetDocumentType when creating a new BigExcelWriter. The default value is now Workbook (.xlsx).
- Improved documentation comments. Also published in https://fenase.github.io/BigExcelCreator/api/BigExcelCreator.html

## 3.1 (and 2.3.2024.30215)
### Added
- Overloads for WriteNumberCell and WriteNumberRow to accept different numeric types in addition to float.

## 3.0
This version uses the new 3.* version of DocumentFormat.OpenXml and it's not compatible with 2.*
If you're using another package that require DocumentFormat.OpenXml 2.*, consider using version 2.3 of this package.

### Added
- .Net8.0 build target
### Removed
- .NetStandard1.3 build target

## 2.3
### Removed
- Finally removed method AddValidator (marked as obsolete since before version 1)
### Added
- Integer and decimal data validation
- .Net6.0 build target
### Changed
- Throwing more specific exceptions instead of just throwing InvalidOperacionException for everything
- Dependency update: DocumentFormat.OpenXml 2.18.0 -> 2.20.0

## 2.2
### Added
- Show or hide, print or not, Gridlines and headings
### Changed
- Bumped dependencies version to current latest since the reason to lower it no longer applies.

## 2.1
### Changed
- Lowered minimum required version of DocumentFormat.OpenXml. It is still recommended to use the latest version when possible.
### Added
- Ability to merge cells

## 2.0
### Changed
- Renamed class BigExcelWritter to BigExcelWriter.
  Sorry for the typo.
### Added
- Conditional formatting
    - By formula
    - By value (Cell Is)
    - Duplicated values

## 1.1
### Added
- Text cells can now be written as shared strings instead of as value. This should reduce the final file's size when the same text is repeated across sheets

## 1.0
- First version considered to be "stable".
- Moved repository to GitHub (previously hosted on Azure DevOps)
### Changed
- Renamed `WriteTextCell<int>` to `WriteNumberCell<int>`. `WriteTextCell<string>` is still in use.

## 1.0.265
### Added
- Hide rows and columns
- Write formula to cell

## 0.2022.262
### Added
- Create autofilter
- Ranges are now validated

## 0.2022.261
### Added
- Add comments to cells

## 0.2022.256
### Added
- Styling and formatting

## 0.2022.253
- Initial version
