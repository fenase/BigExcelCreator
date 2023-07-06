# Changelog

## 2.3
### Removed
- Finally removed method AddValidator (marked as obsolete since before version 1)
### Added
- .Net6.0 build target
### Changed
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
