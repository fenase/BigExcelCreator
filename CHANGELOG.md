# Changelog

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
- Moved repository to GitHub (previously was on Azure DevOps)
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
