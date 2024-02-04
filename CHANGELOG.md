# Changelog

All notable changes to ["Au3LibreOffice"](https://github.com/mlipok/Au3LibreOffice/tree/main) SDK/API will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
This project also adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Go to [legend](#legend---types-of-changes) for further information about the types of changes.

## Releases

|    Version       |    Changes                   |    Download                            |     Released   |    Compare on GitHub       |
|:-----------------|:----------------------------:|:--------------------------------------:|:--------------:|:---------------------------|
|    **v0.10.0**   | [Change Log](##[0.10.0])     | [v0.10.0][v0.10.0]                     | _Unreleased_   | [Compare][v0.10.0-Compare] |
|    **v0.9.1**    | [Change Log](##[0.9.1])      | [v0.9.1][v0.9.1]                       | 2023-10-28     | [Compare][v0.9.1-Compare]  |
|    **v0.9.0**    | [Change Log](##[0.9.0])      | [v0.9.0][v0.9.0]                       | 2023-10-28     | [Compare][v0.9.0-Compare]  |
|    **v0.0.0.3**  | [Change Log](##[0.0.0.3])    | [v0.0.0.3][v0.0.0.3]                   | 2023-08-10     | [Compare][v0.0.0.3-Compare]|
|    **v0.0.0.2**  | [Change Log](##[0.0.0.2])    | [v0.0.0.2][v0.0.0.2]                   | 2023-07-16     | [Compare][v0.0.0.2-Compare]|
|    **v0.0.0.1**  | [Change Log](##[0.0.0.1])    | [v0.0.0.1][v0.0.0.1]                   | 2023-07-02     |                            |

## [0.10.0] - 2024-

### Project

- Added logo to ReadMe. @mLipok
- Filled in ReadMe. @mLipok
- Formatted Changelog

> [!NOTE]
> **LibreOfficeUDF**

### Added

- Central Constants File
	- LibreOffice_Constants.au3
- Central UDF File for all components (@mLipok)
	- LibreOffice.au3

### Changed

- All Internal Error Constants from $__LOW_STATUS_ or $__LOC_STATUS_ To $__LO_STATUS_

> [!NOTE]
> **LibreOfficeCalc**

### Added

- Main Calc File
	- LibreOfficeCalc.au3
- Individual Calc Element Files
	- LibreOfficeCalc_Cell.au3
	- LibreOfficeCalc_CellStyle.au3
	- LibreOfficeCalc_Constants.au3
	- LibreOfficeCalc_Doc.au3
	- LibreOfficeCalc_Font.au3
	- LibreOfficeCalc_Helper.au3
	- LibreOfficeCalc_Internal.au3
	- LibreOfficeCalc_Range.au3
	- LibreOfficeCalc_Sheet.au3
- Cell/Cell Range Formatting Functions and Examples
	- _LOCalc_CellBackColor
	- _LOCalc_CellBorderColor
	- _LOCalc_CellBorderPadding
	- _LOCalc_CellBorderStyle
	- _LOCalc_CellBorderWidth
	- _LOCalc_CellEffect
	- _LOCalc_CellFont
	- _LOCalc_CellFontColor
	- _LOCalc_CellFormula
	- _LOCalc_CellGetType
	- _LOCalc_CellNumberFormat
	- _LOCalc_CellOverline
	- _LOCalc_CellProtection
	- _LOCalc_CellShadow
	- _LOCalc_CellStrikeOut
	- _LOCalc_CellString
	- _LOCalc_CellTextAlign
	- _LOCalc_CellTextOrient
	- _LOCalc_CellTextProperties
	- _LOCalc_CellUnderline
	- _LOCalc_CellValue
- Cell/Cell Range Functions and Examples
	- _LOCalc_RangeClearContents
	- _LOCalc_RangeColumnDelete
	- _LOCalc_RangeColumnGetName
	- _LOCalc_RangeColumnGetObjByName
	- _LOCalc_RangeColumnGetObjByPosition
	- _LOCalc_RangeColumnInsert
	- _LOCalc_RangeColumnPageBreak
	- _LOCalc_RangeColumnsGetCount
	- _LOCalc_RangeColumnVisible
	- _LOCalc_RangeColumnWidth
	- _LOCalc_RangeCopyMove
	- _LOCalc_RangeData
	- _LOCalc_RangeDelete
	- _LOCalc_RangeFormula
	- _LOCalc_RangeGetCellByName
	- _LOCalc_RangeGetCellByPosition
	- _LOCalc_RangeInsert
	- _LOCalc_RangeNumbers
	- _LOCalc_RangeQueryColumnDiff
	- _LOCalc_RangeQueryContents
	- _LOCalc_RangeQueryDependents
	- _LOCalc_RangeQueryEmpty
	- _LOCalc_RangeQueryFormula
	- _LOCalc_RangeQueryIntersection
	- _LOCalc_RangeQueryPrecedents
	- _LOCalc_RangeQueryRowDiff
	- _LOCalc_RangeQueryVisible
	- _LOCalc_RangeRowDelete
	- _LOCalc_RangeRowGetObjByPosition
	- _LOCalc_RangeRowHeight
	- _LOCalc_RangeRowInsert
	- _LOCalc_RangeRowPageBreak
	- _LOCalc_RangeRowsGetCount
	- _LOCalc_RangeRowVisible
- Cell Style Formatting Functions and Examples
	- _LOCalc_CellStyleBackColor
	- _LOCalc_CellStyleBorderColor
	- _LOCalc_CellStyleBorderPadding
	- _LOCalc_CellStyleBorderStyle
	- _LOCalc_CellStyleBorderWidth
	- _LOCalc_CellStyleCreate
	- _LOCalc_CellStyleDelete
	- _LOCalc_CellStyleEffect
	- _LOCalc_CellStyleExists
	- _LOCalc_CellStyleFont
	- _LOCalc_CellStyleFontColor
	- _LOCalc_CellStyleGetObj
	- _LOCalc_CellStyleNumberFormat
	- _LOCalc_CellStyleOrganizer
	- _LOCalc_CellStyleOverline
	- _LOCalc_CellStyleProtection
	- _LOCalc_CellStyleSet
	- _LOCalc_CellStylesGetNames
	- _LOCalc_CellStyleShadow
	- _LOCalc_CellStyleStrikeOut
	- _LOCalc_CellStyleTextAlign
	- _LOCalc_CellStyleTextOrient
	- _LOCalc_CellStyleTextProperties
	- _LOCalc_CellStyleUnderline
- Document Functions and Examples
	- _LOCalc_DocClose
	- _LOCalc_DocConnect
	- _LOCalc_DocCreate
	- _LOCalc_DocExport
	- _LOCalc_DocGetName
	- _LOCalc_DocGetPath
	- _LOCalc_DocHasPath
	- _LOCalc_DocHasSheetName
	- _LOCalc_DocIsActive
	- _LOCalc_DocIsModified
	- _LOCalc_DocIsReadOnly
	- _LOCalc_DocMaximize
	- _LOCalc_DocMinimize
	- _LOCalc_DocOpen
	- _LOCalc_DocPosAndSize
	- _LOCalc_DocRedo
	- _LOCalc_DocRedoClear
	- _LOCalc_DocRedoCurActionTitle
	- _LOCalc_DocRedoGetAllActionTitles
	- _LOCalc_DocRedoIsPossible
	- _LOCalc_DocSave
	- _LOCalc_DocSaveAs
	- _LOCalc_DocToFront
	- _LOCalc_DocUndo
	- _LOCalc_DocUndoActionBegin
	- _LOCalc_DocUndoActionEnd
	- _LOCalc_DocUndoClear
	- _LOCalc_DocUndoCurActionTitle
	- _LOCalc_DocUndoGetAllActionTitles
	- _LOCalc_DocUndoIsPossible
	- _LOCalc_DocUndoReset
	- _LOCalc_DocVisible
	- _LOCalc_DocZoom
- Font Query Functions
	- _LOCalc_FontExists
	- _LOCalc_FontsList
- Helper Functions
	- _LOCalc_ComError_UserFunction
	- _LOCalc_ConvertColorFromLong
	- _LOCalc_ConvertColorToLong
	- _LOCalc_ConvertFromMicrometer
	- _LOCalc_ConvertToMicrometer
	- _LOCalc_FormatKeyCreate
	- _LOCalc_FormatKeyDelete
	- _LOCalc_FormatKeyExists
	- _LOCalc_FormatKeyGetStandard
	- _LOCalc_FormatKeyGetString
	- _LOCalc_FormatKeyList
	- _LOCalc_PathConvert
	- _LOCalc_VersionGet
- Internal Functions
	- __LOCalc_AddTo1DArray
	- __LOCalc_ArrayFill
	- __LOCalc_CellBackColor
	- __LOCalc_CellBorder
	- __LOCalc_CellBorderPadding
	- __LOCalc_CellEffect
	- __LOCalc_CellFont
	- __LOCalc_CellFontColor
	- __LOCalc_CellNumberFormat
	- __LOCalc_CellOverLine
	- __LOCalc_CellProtection
	- __LOCalc_CellShadow
	- __LOCalc_CellStrikeOut
	- __LOCalc_CellStyleBorder
	- __LOCalc_CellTextAlign
	- __LOCalc_CellTextOrient
	- __LOCalc_CellTextProperties
	- __LOCalc_CellUnderLine
	- __LOCalc_CreateStruct
	- __LOCalc_FilterNameGet
	- __LOCalc_InternalComErrorHandler
	- __LOCalc_IntIsBetween
	- __LOCalc_NumIsBetween
	- __LOCalc_SetPropertyValue
	- __LOCalc_UnitConvert
	- __LOCalc_VarsAreDefault
	- __LOCalc_VarsAreNull
	- __LOCalc_VersionCheck
- Sheet Functions and Examples
	- _LOCalc_SheetActivate
	- _LOCalc_SheetAdd
	- _LOCalc_SheetCopy
	- _LOCalc_SheetGetActive
	- _LOCalc_SheetGetObjByName
	- _LOCalc_SheetIsActive
	- _LOCalc_SheetIsProtected
	- _LOCalc_SheetMove
	- _LOCalc_SheetName
	- _LOCalc_SheetProtect
	- _LOCalc_SheetRemove
	- _LOCalc_SheetsGetCount
	- _LOCalc_SheetsGetNames
	- _LOCalc_SheetUnprotect
	- _LOCalc_SheetVisible
- Calc Constants
	- $LOC_BORDERSTYLE_*
	- $LOC_BORDERWIDTH_*
	- $LOC_CELL_ALIGN_HORI_*
	- $LOC_CELL_ALIGN_VERT_*
	- $LOC_CELL_DELETE_MODE_*
	- $LOC_CELL_FLAG_*
	- $LOC_CELL_INSERT_MODE_*
	- $LOC_CELL_ROTATE_REF_*
	- $LOC_CELL_TYPE_*
	- $LOC_COLOR_*
	- $LOC_FORMAT_KEYS_*
	- $LOC_FORMULA_RESULT_TYPE_*
	- $LOC_PATHCONV_*
	- $LOC_POSTURE_*
	- $LOC_RELIEF_*
	- $LOC_SHADOW_*
	- $LOC_STRIKEOUT_*
	- $LOC_TXT_DIR_*
	- $LOC_UNDERLINE_*
	- $LOC_WEIGHT_*
	- $LOC_ZOOMTYPE_*

> [!NOTE]
> **LibreOfficeWriter**

### Added

- More Undo Functions and Examples
	- _LOWriter_DocRedoClear
	- _LOWriter_DocUndoActionBegin
	- _LOWriter_DocUndoActionEnd
	- _LOWriter_DocUndoClear
	- _LOWriter_DocUndoReset
- Shape Functions and Examples
	- _LOWriter_DocHasShapeName
	- _LOWriter_ShapeAreaColor
	- _LOWriter_ShapeAreaGradient
	- _LOWriter_ShapeDelete
	- _LOWriter_ShapeGetAnchor
	- _LOWriter_ShapeGetObjByName
	- _LOWriter_ShapeGetType
	- _LOWriter_ShapeInsert
	- _LOWriter_ShapeLineArrowStyles
	- _LOWriter_ShapeLineProperties
	- _LOWriter_ShapeName
	- _LOWriter_ShapePointsAdd
	- _LOWriter_ShapePointsGetCount
	- _LOWriter_ShapePointsModify
	- _LOWriter_ShapePointsRemove
	- _LOWriter_ShapePosition
	- _LOWriter_ShapeRotateSlant
	- _LOWriter_ShapesGetNames
	- _LOWriter_ShapeTextBox
	- _LOWriter_ShapeTransparency
	- _LOWriter_ShapeTransparencyGradient
	- _LOWriter_ShapeTypePosition
	- _LOWriter_ShapeTypeSize
	- _LOWriter_ShapeWrap
	- _LOWriter_ShapeWrapOptions
	- __LOWriter_CreatePoint
	- __LOWriter_GetShapeName
	- __LOWriter_Shape_CreateArrow
	- __LOWriter_Shape_CreateBasic
	- __LOWriter_Shape_CreateCallout
	- __LOWriter_Shape_CreateFlowchart
	- __LOWriter_Shape_CreateLine
	- __LOWriter_Shape_CreateStars
	- __LOWriter_Shape_CreateSymbol
	- __LOWriter_Shape_GetCustomType
	- __LOWriter_ShapeArrowStyleName
	- __LOWriter_ShapeLineStyleName
	- __LOWriter_ShapePointGetSettings
	- __LOWriter_ShapePointModify
- Shape Point Constants in LibreOfficeWriter_Constants. $LOW_SHAPE_POINT_TYPE_*
- Standard Format Key retrieval function _LOWriter_FormatKeyGetStandard

### Changed

- ReArranged Parameters in PageBreak functions to be more logically ordered. The previous order was:

	> $iBreakType = Null, $iPgNumOffSet = Null, $sPageStyle = Null

	However the functions stated:

	> _Break Type_ must be set **before** _Page Style_ will be able to be set, and _page style_ needs set **before** _$iPgNumOffSet_ can be set.

	Thus the parameters have been rearranged to:

	> $iBreakType = Null, $sPageStyle = Null, $iPgNumOffSet

	Affected functions are:
	- _LOWriter_DirFrmtParPageBreak
	- _LOWriter_ParStylePageBreak
	- __LOWriter_ParPageBreak

- Undo/Redo Action Title retrieval functions to no longer return an error when Undo/Redo Action Titles are unavailable, instead they  return an empty string or array, respectively.
	- _LOWriter_DocRedoCurActionTitle;
	- _LOWriter_DocRedoGetAllActionTitles;
	- _LOWriter_DocUndoGetAllActionTitles;
	- _LOWriter_DocUndoCurActionTitle.
- Minor documentation adjustments.
- Moved Search Descriptor functions from LibreOfficeWriter.au3 to LibreOfficeWriter_Helper.au3.
	- _LOWriter_SearchDescriptorCreate
	- _LOWriter_SearchDescriptorModify
	- _LOWriter_SearchDescriptorSimilarityModify
- Default value of $__LOWCONST_SLEEP_DIV from 15 to 0.
- Sorted Constants in LibreOfficeWriter_Constants alphabetically.
- Renamed $LOW_FIELDADV_TYPE_* Constants to $LOW_FIELD_ADV_TYPE_* to match formatting of other Field Type Constants.
- Renamed $LibreOfficeWriter_DirectFormating.au3 to LibreOfficeWriter_DirectFormatting.au3 (misspelling correction)

### Fixed

- Missing error return value descriptions in:
	- _LOWriter_FootnoteGetAnchor
	- _LOWriter_EndnoteGetAnchor
	- _LOWriter_FrameGetAnchor
	- _LOWriter_ImageGetAnchor
- Missing Includes in LibreOfficeWriter file.
- Ternary Operators missing Parenthesis in responses.
- Export extension was incorrect, "jjpe" --> "jpe"
- Spell Checked the comments and Headers.
- Spell Checked Examples
- Examples on Error now close any documents opened by the example.
- Missing parameters in Null keyword check for __LOWriter_TableBorder, causing errors in setting Table Horizontal and Vertical borders singly.

[To Top](##Releases)

## [0.9.1] - 2023-10-28

> [!NOTE]
> **LibreOfficeWriter**

### Changed

- Minor documentation adjustments.

[To Top](##Releases)

## [0.9.0] - 2023-10-28

> [!NOTE]
> **LibreOfficeWriter**

### Added

- directive for Au3Check to each UDF file branch. (@mLipok)
- Image functions. (@donnyh13)
	- _ImageAreaColor,
	- _ImageAreaGradient,
	- _ImageAreaTransparency,
	- _ImageAreaTransparencyGradient,
	- _ImageBorderColor,
	- _ImageBorderPadding,
	- _ImageBorderStyle,
	- _ImageBorderWidth,
	- _ImageColorAdjust,
	- _ImageCrop,
	- _ImageDelete,
	- _ImageGetAnchor,
	- _ImageGetObjByName,
	- _ImageHyperlink,
	- _ImageInsert,
	- _ImageModify,
	- _ImageOptions,
	- _ImageOptionsName,
	- _ImageReplace,
	- _ImagesGetNames,
	- _ImageShadow,
	- _ImageSize,
	- _ImageTransparency,
	- _ImageTypePosition,
	- _ImageTypeSize,
	- _ImageWrap,
	- _ImageWrapOptions

### Changed

- Original LibreOffice UDF file split into individual elements, per specific usages. (@mLipok.)
	- LibreOfficeWriter_Cell,
	- LibreOfficeWriter_Char,
	- LibreOfficeWriter_Constants,
	- LibreOfficeWriter_Cursor,
	- LibreOfficeWriter_DirectFormatting,
	- LibreOfficeWriter_Doc,
	- LibreOfficeWriter_Field,
	- LibreOfficeWriter_Font,
	- LibreOfficeWriter_FootEndNotes,
	- LibreOfficeWriter_Frame,
	- LibreOfficeWriter_Helper,
	- LibreOfficeWriter_Images,
	- LibreOfficeWriter_Internal,
	- LibreOfficeWriter_Num,
	- LibreOfficeWriter_Page,
	- LibreOfficeWriter_Par,
	- LibreOfficeWriter_Shapes,
	- LibreOfficeWriter_Table.
- Renamed Constants:
	- $LOW_PAPER_PORTRAIT --> $LOW_PAPER_ORIENT_PORTRAIT
	- $LOW_PAPER_LANDSCAPE --> $LOW_PAPER_ORIENT_LANDSCAPE
- Examples layout and error checking cleaned up. (@mLipok.)
- Examples now work from their separate folder. (@mLipok.)
- Major editing of Header layout for every function. As well as several typo corrections and wordiness. (@mLipok) & (@donnyh13.)
- Constants are now all located in the separate Constants files. (@mLipok.)
- All files re-processed with TIDY. (@mLipok)

### Fixed

- Errors caused by residual function calls from filling in "Related" section of the header. (@mLipok.)
	- _LOWriter_DocFindAllInRange,
	- _LOWriter_DocGenPropTemplate.

### Project

- @mLipok and @donnyh13 began jointly working on this project. — Thanks to @mLipok for his tireless work cleaning up many things in this UDF.
- All Constants descriptions are moved to the Constants file. (@donnyh13)
- Added documentation and improved CHANGELOG.md. (@donnyh13)
- Added bug_report.md. (@donnyh13)
- Added feature_request.md. (@donnyh13)
- Added PULL_REQUEST_TEMPLATE.md. (@donnyh13)
- Added CODE_OF_CONDUCT.md (@donnyh13)
- Added CONTRIBUTING.md. (@donnyh13)

Thanks @danp2 and @Sven-Seyfert. All above mentioned MD documents was based on adequate documents from <https://github.com/Danp2/au3WebDriver>.

[To Top](##Releases)

## [0.0.0.3] - 2023-08-10

> [!NOTE]
> **LibreOfficeWriter**

### Added

- Paragraph Object functions which allows the ability to copy and paste content without using the clipboard quickly. Thanks to user @Heiko for inspiration.
	- _ParObjSelect,
	- _ParObjCopy,
	- _ParObjPaste,
	- _ParObjDelete.
- _DocExecuteDispatch function, adding some shortcuts to certain commands, such as select all, copy/paste content to/from clipboard.
- _DocConvertTextToTable.
- _DocConvertTableToText.
- examples for the new functions.
- "Related" functions section for most function headers.
- Warning to _ShapesGetNames, about Images inserted in a document also being called "TextFrames".
- processing error check.
	- _DocInsertString,
	- _DocInsertControlChar.
- a missing Datatype for possible Cursor data position types, $LOW_CURDATA_HEADER_FOOTER, previously attempting to insert a content while the insertion point was located in a Header/Footer would have failed. — Thanks to user @Heiko for helping me locate this error. The affected functions are:
	- _CursorGetDataType.
	- _DocCreateTextCursor,
	- _EndnoteInsert,
	- _FootnoteInsert,
	- _TableInsert,

### Changed

- Renamed Name functions and examples for consistency.
	- `_DocListTableNames` --> `_TablesGetNames`
	- `_FramesListNames` --> `_FramesGetNames`
	- `_ShapesListNames` --> `_ShapesGetNames` 	
	- `_ParGetObjects` --> `_ParObjCreateList`.
	- `_ParSectionsGet` --> `_ParObjSectionsGet`.
	- `_TableGetByCursor` --> `_TableGetObjByCursor`.
	- `_TableGetByName` --> `_TableGetObjByName`.
	- `_TableGetCellByCursor` --> `_TableGetCellObjByCursor`.
	- `_TableGetCellByName` --> `_TableGetCellObjByName`.
	- `_TableGetCellByPosition` --> `_TableGetCellObjByPosition`.
- Removed "IsCollapsed" check and error from _DocGetString.
- _FramesListNames to have an option to search for Frames listed under shapes.
- _ShapesGetNames, Corrected an error that could occur if images are present.

### Fixed

- An error where a COM error would be produced when attempting to insert a string or control character in certain data types. — Thanks to user @Heiko for helping me locate this error.
	- _DocInsertControlChar,
	- _DocInsertString.

[To Top](##Releases)

## [0.0.0.2] - 2023-07-16

> [!NOTE]
> **LibreOfficeWriter**

### Added

- UDF version number in the UDF Header.

### Changed

- _DocReplaceAllInRange to have two methods of performing a Regular Expression find and replace.
- Removed the if/else block in $atFindFormat parameter checking.
	- _DocReplaceAll,
	- _DocReplaceAllInRange,
	- _DocFindNext,
	- _DocFindAll,
	- _DocFindAllInRange.

### Fixed

- Method for skipping $atFindFormat and $atReplaceFormat, now uses an empty array called in each parameter to skip.
	- _DocReplaceAll,
	- _DocReplaceAllInRange.

### Project

- Updated function documentation to reflect the changes.

[To Top](##Releases)

## [0.0.0.1] - 2023-07-02

> [!NOTE]
> **LibreOfficeWriter**

### Added

- Initial UDF Release.

[To Top](##Releases)

---

### Legend - Types of changes

- `Added` for new features.
- `Changed` for changes in existing functionality.
- `Deprecated` for soon-to-be removed features.
- `Fixed` for any bug fixes.
- `Removed` for now removed features.
- `Security` in case of vulnerabilities.
- `Project` for documentation or overall project improvements.

[To the top](#changelog)

---

[v0.10.0-Compare]:	https://github.com/mlipok/Au3LibreOffice/compare/0.9.1...main
[v0.9.1-Compare]:	https://github.com/mlipok/Au3LibreOffice/compare/v0.9.0...0.9.1
[v0.9.0-Compare]:	https://github.com/mlipok/Au3LibreOffice/compare/v0.0.0.3...v0.9.0
[v0.0.0.3-Compare]:	https://github.com/donnyh13/Au3LibreOffice/compare/v0.0.0.2...v0.0.0.3
[v0.0.0.2-Compare]:	https://github.com/donnyh13/Au3LibreOffice/compare/v0.0.0.1...v0.0.0.2

[v0.10.0]:	https://github.com/mlipok/Au3LibreOffice
[v0.9.1]:	https://github.com/mlipok/Au3LibreOffice/releases/tag/0.9.1
[v0.9.0]:	https://github.com/mlipok/Au3LibreOffice/releases/tag/v0.9.0
[v0.0.0.3]:	https://github.com/mlipok/Au3LibreOffice/releases/tag/v0.0.0.3
[v0.0.0.2]:	https://github.com/donnyh13/Au3LibreOffice/releases/tag/v0.0.0.2
[v0.0.0.1]:	https://github.com/donnyh13/Au3LibreOffice/releases/tag/v0.0.0.1
