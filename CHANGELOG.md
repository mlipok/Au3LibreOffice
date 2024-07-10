# Changelog

All notable changes to ["Au3LibreOffice"](https://github.com/mlipok/Au3LibreOffice/tree/main) SDK/API will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
This project also adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Go to [legend](#legend---types-of-changes) for further information about the types of changes.

## Releases

|    Version       |    Changes                         |    Download                 |     Released   |    Compare on GitHub       |
|:-----------------|:----------------------------------:|:---------------------------:|:--------------:|:---------------------------|
|    **v0.10.0**   | [Change Log](#0100---2024-04-)     | [v0.10.0][v0.10.0]          | _Unreleased_   | [Compare][v0.10.0-Compare] |
|    **v0.9.1**    | [Change Log](#091---2023-10-28)    | [v0.9.1][v0.9.1]            | 2023-10-28     | [Compare][v0.9.1-Compare]  |
|    **v0.9.0**    | [Change Log](#090---2023-10-28)    | [v0.9.0][v0.9.0]            | 2023-10-28     | [Compare][v0.9.0-Compare]  |
|    **v0.0.0.3**  | [Change Log](#0003---2023-08-10)   | [v0.0.0.3][v0.0.0.3]        | 2023-08-10     | [Compare][v0.0.0.3-Compare]|
|    **v0.0.0.2**  | [Change Log](#0002---2023-07-16)   | [v0.0.0.2][v0.0.0.2]        | 2023-07-16     | [Compare][v0.0.0.2-Compare]|
|    **v0.0.0.1**  | [Change Log](#0001---2023-07-02)   | [v0.0.0.1][v0.0.0.1]        | 2023-07-02     |                            |

## [0.10.0] - 2024-04-

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
- Attempted to standardize $__LO_STATUS_INIT_ERROR and $__LO_STATUS_PROCESSING_ERROR usage throughout functions.
- Removed Error returns listed in Function Headers that no longer existed.
- Added missing error values and corrected wrong error values listed in the headers.

### Fixed

- Align Parameters, Error/Return values, Remarks, and Related, to the same position.

### Removed

- "Note" from Remarks section in Header. (@mLipok)
- Double spaces from Headers.
- Tabs from headers, replaced with spaces.
- Manual line breaks from headers.

> [!NOTE]
> **LibreOfficeCalc**

### Added

- Main Calc File
	- LibreOfficeCalc.au3
- Individual Calc Element Files
	- LibreOfficeCalc_Cell.au3
	- LibreOfficeCalc_CellStyle.au3
	- LibreOfficeCalc_Comments.au3
	- LibreOfficeCalc_Constants.au3
	- LibreOfficeCalc_Cursor.au3
	- LibreOfficeCalc_Doc.au3
	- LibreOfficeCalc_Field.au3
	- LibreOfficeCalc_Font.au3
	- LibreOfficeCalc_Helper.au3
	- LibreOfficeCalc_Internal.au3
	- LibreOfficeCalc_Page.au3
	- LibreOfficeCalc_Range.au3
	- LibreOfficeCalc_Sheet.au3
- Cell/Cell Range Formatting Functions and Examples
	- _LOCalc_CellBackColor
	- _LOCalc_CellBorderColor
	- _LOCalc_CellBorderPadding
	- _LOCalc_CellBorderStyle
	- _LOCalc_CellBorderWidth
	- _LOCalc_CellCreateTextCursor
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
	- _LOCalc_RangeAutoOutline
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
	- _LOCalc_RangeCompute
	- _LOCalc_RangeCopyMove
	- _LOCalc_RangeCreateCursor
	- _LOCalc_RangeData
	- _LOCalc_RangeDatabaseAdd
	- _LOCalc_RangeDatabaseDelete
	- _LOCalc_RangeDatabaseGetNames
	- _LOCalc_RangeDatabaseGetObjByName
	- _LOCalc_RangeDatabaseHasByName
	- _LOCalc_RangeDatabaseModify
	- _LOCalc_RangeDelete
	- _LOCalc_RangeDetail
	- _LOCalc_RangeFill
	- _LOCalc_RangeFillSeries
	- _LOCalc_RangeFilter
	- _LOCalc_RangeFilterClear
	- _LOCalc_RangeFindAll
	- _LOCalc_RangeFindNext
	- _LOCalc_RangeFormula
	- _LOCalc_RangeGetAddressAsName
	- _LOCalc_RangeGetAddressAsPosition
	- _LOCalc_RangeGetCellByName
	- _LOCalc_RangeGetCellByPosition
	- _LOCalc_RangeGetSheet
	- _LOCalc_RangeGroup
	- _LOCalc_RangeInsert
	- _LOCalc_RangeIsMerged
	- _LOCalc_RangeMerge
	- _LOCalc_RangeNamedAdd
	- _LOCalc_RangeNamedChangeScope
	- _LOCalc_RangeNamedDelete
	- _LOCalc_RangeNamedGetNames
	- _LOCalc_RangeNamedGetObjByName
	- _LOCalc_RangeNamedHasByName
	- _LOCalc_RangeNamedModify
	- _LOCalc_RangeNumbers
	- _LOCalc_RangeOutlineClearAll
	- _LOCalc_RangeOutlineShow
	- _LOCalc_RangeQueryColumnDiff
	- _LOCalc_RangeQueryContents
	- _LOCalc_RangeQueryDependents
	- _LOCalc_RangeQueryEmpty
	- _LOCalc_RangeQueryFormula
	- _LOCalc_RangeQueryIntersection
	- _LOCalc_RangeQueryPrecedents
	- _LOCalc_RangeQueryRowDiff
	- _LOCalc_RangeQueryVisible
	- _LOCalc_RangeReplace
	- _LOCalc_RangeReplaceAll
	- _LOCalc_RangeRowDelete
	- _LOCalc_RangeRowGetObjByPosition
	- _LOCalc_RangeRowHeight
	- _LOCalc_RangeRowInsert
	- _LOCalc_RangeRowPageBreak
	- _LOCalc_RangeRowsGetCount
	- _LOCalc_RangeRowVisible
	- _LOCalc_RangeSort
	- _LOCalc_RangeSortAlt
	- _LOCalc_RangeValidation
	- _LOCalc_RangeValidationSettings
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
- Comment Functions and Examples.
	- _LOCalc_CommentAdd
	- _LOCalc_CommentAreaColor
	- _LOCalc_CommentAreaFillStyle
	- _LOCalc_CommentAreaGradient
	- _LOCalc_CommentAreaShadow
	- _LOCalc_CommentAreaTransparency
	- _LOCalc_CommentAreaTransparencyGradient
	- _LOCalc_CommentCallout
	- _LOCalc_CommentCreateTextCursor
	- _LOCalc_CommentDelete
	- _LOCalc_CommentGetCell
	- _LOCalc_CommentGetLastEdit
	- _LOCalc_CommentGetObjByCell
	- _LOCalc_CommentGetObjByIndex
	- _LOCalc_CommentLineArrowStyles
	- _LOCalc_CommentLineProperties
	- _LOCalc_CommentPosition
	- _LOCalc_CommentRotate
	- _LOCalc_CommentsGetCount
	- _LOCalc_CommentsGetList
	- _LOCalc_CommentSize
	- _LOCalc_CommentText
	- _LOCalc_CommentTextAnchor
	- _LOCalc_CommentTextAnimation
	- _LOCalc_CommentTextColumns
	- _LOCalc_CommentTextSettings
	- _LOCalc_CommentVisible
- Cursor Functions and Examples
	- _LOCalc_SheetCursorMove
	- _LOCalc_TextCursorCharPosition
	- _LOCalc_TextCursorCharSpacing
	- _LOCalc_TextCursorEffect
	- _LOCalc_TextCursorFont
	- _LOCalc_TextCursorFontColor
	- _LOCalc_TextCursorGetString
	- _LOCalc_TextCursorGoToRange
	- _LOCalc_TextCursorInsertString
	- _LOCalc_TextCursorIsCollapsed
	- _LOCalc_TextCursorMove
	- _LOCalc_TextCursorOverline
	- _LOCalc_TextCursorParObjCreateList
	- _LOCalc_TextCursorParObjSectionsGet
	- _LOCalc_TextCursorStrikeOut
	- _LOCalc_TextCursorUnderline
- Document Functions and Examples
	- _LOCalc_DocClose
	- _LOCalc_DocColumnsRowsAreFrozen
	- _LOCalc_DocColumnsRowsFreeze
	- _LOCalc_DocConnect
	- _LOCalc_DocCreate
	- _LOCalc_DocEnumPrinters
	- _LOCalc_DocEnumPrintersAlt
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
	- _LOCalc_DocPrint
	- _LOCalc_DocRedo
	- _LOCalc_DocRedoClear
	- _LOCalc_DocRedoCurActionTitle
	- _LOCalc_DocRedoGetAllActionTitles
	- _LOCalc_DocRedoIsPossible
	- _LOCalc_DocSave
	- _LOCalc_DocSaveAs
	- _LOCalc_DocSelectionCopy
	- _LOCalc_DocSelectionGet
	- _LOCalc_DocSelectionPaste
	- _LOCalc_DocSelectionSet
	- _LOCalc_DocSelectionSetMulti
	- _LOCalc_DocToFront
	- _LOCalc_DocUndo
	- _LOCalc_DocUndoActionBegin
	- _LOCalc_DocUndoActionEnd
	- _LOCalc_DocUndoClear
	- _LOCalc_DocUndoCurActionTitle
	- _LOCalc_DocUndoGetAllActionTitles
	- _LOCalc_DocUndoIsPossible
	- _LOCalc_DocUndoReset
	- _LOCalc_DocViewDisplaySettings
	- _LOCalc_DocViewWindowSettings
	- _LOCalc_DocVisible
	- _LOCalc_DocWindowFirstColumn
	- _LOCalc_DocWindowFirstRow
	- _LOCalc_DocWindowIsSplit
	- _LOCalc_DocWindowSplit
	- _LOCalc_DocWindowVisibleRange
	- _LOCalc_DocZoom
- Field Functions
	- _LOCalc_FieldCurrentDisplayGet
	- _LOCalc_FieldDateTimeInsert
	- _LOCalc_FieldDelete
	- _LOCalc_FieldFileNameInsert
	- _LOCalc_FieldGetAnchor
	- _LOCalc_FieldHyperlinkInsert
	- _LOCalc_FieldHyperlinkModify
	- _LOCalc_FieldPageCountInsert
	- _LOCalc_FieldPageNumberInsert
	- _LOCalc_FieldsGetList
	- _LOCalc_FieldSheetNameInsert
	- _LOCalc_FieldTitleInsert
- Font Query Functions
	- _LOCalc_FontExists
	- _LOCalc_FontsList
- Helper Functions
	- _LOCalc_ComError_UserFunction
	- _LOCalc_ConvertColorFromLong
	- _LOCalc_ConvertColorToLong
	- _LOCalc_ConvertFromMicrometer
	- _LOCalc_ConvertToMicrometer
	- _LOCalc_FilterDescriptorCreate
	- _LOCalc_FilterDescriptorModify
	- _LOCalc_FilterFieldCreate
	- _LOCalc_FilterFieldModify
	- _LOCalc_FormatKeyCreate
	- _LOCalc_FormatKeyDelete
	- _LOCalc_FormatKeyExists
	- _LOCalc_FormatKeyGetStandard
	- _LOCalc_FormatKeyGetString
	- _LOCalc_FormatKeyList
	- _LOCalc_PathConvert
	- _LOCalc_SearchDescriptorCreate
	- _LOCalc_SearchDescriptorModify
	- _LOCalc_SearchDescriptorSimilarityModify
	- _LOCalc_SortFieldCreate
	- _LOCalc_SortFieldModify
	- _LOCalc_VersionGet
- Internal Functions
	- __LOCalc_AddTo1DArray
	- __LOCalc_ArrayFill
	- __LOCalc_CellAddressIsSame
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
	- __LOCalc_CharPosition
	- __LOCalc_CharSpacing
	- __LOCalc_CommentAreaShadowModify
	- __LOCalc_CommentArrowStyleName
	- __LOCalc_CommentGetObjByCell
	- __LOCalc_CommentLineStyleName
	- __LOCalc_CreateStruct
	- __LOCalc_FieldGetObj
	- __LOCalc_FieldTypeServices
	- __LOCalc_FilterNameGet
	- __LOCalc_GradientNameInsert
	- __LOCalc_GradientPresets
	- __LOCalc_Internal_CursorGetType
	- __LOCalc_InternalComErrorHandler
	- __LOCalc_IntIsBetween
	- __LOCalc_NamedRangeGetScopeObj
	- __LOCalc_NumIsBetween
	- __LOCalc_PageStyleBorder
	- __LOCalc_PageStyleFooterBorder
	- __LOCalc_PageStyleHeaderBorder
	- __LOCalc_RangeAddressIsSame
	- __LOCalc_SetPropertyValue
	- __LOCalc_SheetCursorMove
	- __LOCalc_TextCursorMove
	- __LOCalc_TransparencyGradientConvert
	- __LOCalc_TransparencyGradientNameInsert
	- __LOCalc_UnitConvert
	- __LOCalc_VarsAreDefault
	- __LOCalc_VarsAreNull
	- __LOCalc_VersionCheck
- Page Style Functions and Examples
	- _LOCalc_PageStyleAreaColor
	- _LOCalc_PageStyleBorderColor
	- _LOCalc_PageStyleBorderPadding
	- _LOCalc_PageStyleBorderStyle
	- _LOCalc_PageStyleBorderWidth
	- _LOCalc_PageStyleCreate
	- _LOCalc_PageStyleDelete
	- _LOCalc_PageStyleExists
	- _LOCalc_PageStyleFooter
	- _LOCalc_PageStyleFooterAreaColor
	- _LOCalc_PageStyleFooterBorderColor
	- _LOCalc_PageStyleFooterBorderPadding
	- _LOCalc_PageStyleFooterBorderStyle
	- _LOCalc_PageStyleFooterBorderWidth
	- _LOCalc_PageStyleFooterCreateTextCursor
	- _LOCalc_PageStyleFooterObj
	- _LOCalc_PageStyleFooterShadow
	- _LOCalc_PageStyleGetObj
	- _LOCalc_PageStyleHeader
	- _LOCalc_PageStyleHeaderAreaColor
	- _LOCalc_PageStyleHeaderBorderColor
	- _LOCalc_PageStyleHeaderBorderPadding
	- _LOCalc_PageStyleHeaderBorderStyle
	- _LOCalc_PageStyleHeaderBorderWidth
	- _LOCalc_PageStyleHeaderCreateTextCursor
	- _LOCalc_PageStyleHeaderObj
	- _LOCalc_PageStyleHeaderShadow
	- _LOCalc_PageStyleLayout
	- _LOCalc_PageStyleMargins
	- _LOCalc_PageStyleOrganizer
	- _LOCalc_PageStylePaperFormat
	- _LOCalc_PageStyleSet
	- _LOCalc_PageStylesGetNames
	- _LOCalc_PageStyleShadow
	- _LOCalc_PageStyleSheetPageOrder
	- _LOCalc_PageStyleSheetPrint
	- _LOCalc_PageStyleSheetScale
- Sheet Functions and Examples
	- _LOCalc_SheetActivate
	- _LOCalc_SheetAdd
	- _LOCalc_SheetCopy
	- _LOCalc_SheetCreateCursor
	- _LOCalc_SheetDetectiveClear
	- _LOCalc_SheetDetectiveDependent
	- _LOCalc_SheetDetectiveInvalidData
	- _LOCalc_SheetDetectivePrecedent
	- _LOCalc_SheetDetectiveTraceError
	- _LOCalc_SheetGetActive
	- _LOCalc_SheetGetObjByName
	- _LOCalc_SheetGetObjByPosition
	- _LOCalc_SheetImport
	- _LOCalc_SheetIsActive
	- _LOCalc_SheetIsProtected
	- _LOCalc_SheetLink
	- _LOCalc_SheetLinkModify
	- _LOCalc_SheetMove
	- _LOCalc_SheetName
	- _LOCalc_SheetPrintColumnsRepeat
	- _LOCalc_SheetPrintRangeModify
	- _LOCalc_SheetPrintRowsRepeat
	- _LOCalc_SheetProtect
	- _LOCalc_SheetRemove
	- _LOCalc_SheetsGetCount
	- _LOCalc_SheetsGetNames
	- _LOCalc_SheetTabColor
	- _LOCalc_SheetUnprotect
	- _LOCalc_SheetVisible
- Calc Constants
	- $__LOCCONST_FILL_STYLE_*
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
	- $LOC_COMMENT_ANCHOR_*
	- $LOC_COMMENT_ANIMATION_DIR_*
	- $LOC_COMMENT_ANIMATION_KIND_*
	- $LOC_COMMENT_CALLOUT_EXT_ALIGN_HORI_*
	- $LOC_COMMENT_CALLOUT_EXT_ALIGN_VERT_*
	- $LOC_COMMENT_CALLOUT_EXT_*
	- $LOC_COMMENT_CALLOUT_STYLE_*
	- $LOC_COMMENT_LINE_ARROW_TYPE_*
	- $LOC_COMMENT_LINE_CAP_*
	- $LOC_COMMENT_LINE_JOINT_*
	- $LOC_COMMENT_LINE_STYLE_*
	- $LOC_COMMENT_SHADOW_*
	- $LOC_COMPUTE_*
	- $LOC_CURTYPE_*
	- $LOC_DUPLEX_*
	- $LOC_FIELD_TYPE_*
	- $LOC_FILL_DATE_MODE_*
	- $LOC_FILL_DIR_*
	- $LOC_FILL_MODE_*
	- $LOC_FILTER_CONDITION_*
	- $LOC_FILTER_OPERATOR_*
	- $LOC_FORMAT_KEYS_*
	- $LOC_FORMULA_RESULT_TYPE_*
	- $LOC_GRAD_NAME_*
	- $LOC_GRAD_TYPE_*
	- $LOC_GROUP_ORIENT_*
	- $LOC_NAMED_RANGE_OPT_*
	- $LOC_NUM_STYLE_*
	- $LOC_PAGE_LAYOUT_*
	- $LOC_PAPER_HEIGHT_*
	- $LOC_PAPER_WIDTH_*
	- $LOC_PATHCONV_*
	- $LOC_POSTURE_*
	- $LOC_RELIEF_*
	- $LOC_SCALE_*
	- $LOC_SEARCH_IN_*
	- $LOC_SHADOW_*
	- $LOC_SHEET_LINK_MODE_*
	- $LOC_SHEETCUR_*
	- $LOC_SORT_DATA_TYPE_*
	- $LOC_STRIKEOUT_*
	- $LOC_TEXTCUR_*
	- $LOC_TXT_DIR_*
	- $LOC_UNDERLINE_*
	- $LOC_VALIDATION_COND_*
	- $LOC_VALIDATION_ERROR_*
	- $LOC_VALIDATION_LIST_*
	- $LOC_VALIDATION_TYPE_*
	- $LOC_WEIGHT_*
	- $LOC_ZOOMTYPE_*
- Auto size option to Range Data, Formulas, and Numbers fill functions.
- Retrieve Linked Sheet names only to _LOCalc_SheetsGetNames.

### Fixed

- Removed unused variables and parameters in some functions. Affected functions are as follows:
	- _LOCalc_FormatKeyDelete -- removed internal variable.
- NamedRange names incorrectly reported as invalid in certain functions when the name began with an underscore.
	- _LOCalc_RangeNamedModify
	- _LOCalc_RangeNamedAdd
- _LOCalc_DocOpen now uses a different method for connecting to an already open document, as the previous method was causing errors.

### Changed

- Constant $__LOCCONST_FILL_STYLE_* to $LOC_AREA_FILL_STYLE_*
- __LOCalc_IntIsBetween to accept only a minimum value. Also optimized it.
	- Modified function usage to match changes.

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
- Alpha Removal function __LOWriter_ColorRemoveAlpha.
- _LOWriter_FrameAreaFillStyle.
- _LOWriter_FrameStyleAreaFillStyle
- _LOWriter_ImageAreaFillStyle
- _LOWriter_PageStyleAreaFillStyle
- _LOWriter_PageStyleFooterAreaFillStyle
- _LOWriter_PageStyleHeaderAreaFillStyle
- _LOWriter_ShapeAreaFillStyle

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
- Renamed Parameter $iBorder for Gradient Functions to $iTransitionStart for clarity of which setting it corresponds to in the L.O. UI.
	- _LOWriter_FrameAreaGradient
	- _LOWriter_FrameTransparencyGradient
	- _LOWriter_FrameStyleAreaGradient
	- _LOWriter_FrameStyleTransparencyGradient
	- _LOWriter_ImageAreaGradient
	- _LOWriter_ImageAreaTransparencyGradient
	- _LOWriter_PageStyleAreaGradient
	- _LOWriter_PageStyleFooterAreaGradient
	- _LOWriter_PageStyleFooterTransparencyGradient
	- _LOWriter_PageStyleHeaderAreaGradient
	- _LOWriter_PageStyleHeaderTransparencyGradient
	- _LOWriter_PageStyleTransparencyGradient
	- _LOWriter_ShapeAreaGradient
	- _LOWriter_ShapeTransparencyGradient
- _LOWriter_DocOpen now uses a different method for connecting to an already open document, as the previous method could potentially cause errors.
- Renamed $__LOWCONST_FILL_STYLE_* Constant to $LOW_AREA_FILL_STYLE_*.
- __LOWriter_IntIsBetween to accept only a minimum value. Also optimized it.
	- Modified function usage to match changes.

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
- Wrong error value in _LOWriter_CursorGetStatus for error while determining cursor type.
- Missing Parameters in some headers or Syntax header.
- Missing parameter type checks in some functions.
- Missing ExitLoop in __LOWriter_FieldsGetList causing unnecessary looping.
- Removed unused variables and parameters in some functions. Affected functions are as follows:
	- _LOWriter_DocBookmarkModify -- removed $oDoc parameter.
	- _LOWriter_FieldDelete -- removed $oDoc parameter.
	- _LOWriter_FieldDocInfoEditTimeModify -- removed $oDoc parameter.
	- _LOWriter_FieldSetVarMasterListFields -- removed $oDoc parameter.
	- _LOWriter_DateFormatKeyDelete -- removed internal variable.
	- _LOWriter_FormatKeyDelete -- removed internal variable.
	- _LOWriter_ImageInsert -- removed internal variable, now uses ViewCursor directly to insert an Image.
	- _LOWriter_NumStyleSetLevel -- removed $oDoc parameter.
- _LOWriter_DocHyperlinkInsert -- $bOverwrite parameter was not used in function.
- Transparency causing Color values to be returned that including the Alpha value, causing potentially unexpected results.
- Missing Data Type in _LOWriter_CursorGetDataType example.
- _LOWriter_DocCreateTextCursor would throw an error when creating a Text Cursor at the ViewCursor position.
- When any FieldsGetList functions were supposed to return a single dimension array, a two dimensional array was being returned.
- Incorrect usage of ObjEvent.

### Removed

- $LOW_FIELD_TYPE_URL Constant. -- "com.sun.star.text.TextField.URL" is a Calc-only Field type.

[To Top](#releases)

## [0.9.1] - 2023-10-28

> [!NOTE]
> **LibreOfficeWriter**

### Changed

- Minor documentation adjustments.

[To Top](#releases)

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

[To Top](#releases)

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

[To Top](#releases)

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

[To Top](#releases)

## [0.0.0.1] - 2023-07-02

> [!NOTE]
> **LibreOfficeWriter**

### Added

- Initial UDF Release.

[To Top](#releases)

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
