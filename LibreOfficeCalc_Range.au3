#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Internal.au3"

; Other includes for Calc

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, or applying settings to L.O. Calc Cell Ranges.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_RangeClearContents
; _LOCalc_RangeColumnDelete
; _LOCalc_RangeColumnGetName
; _LOCalc_RangeColumnGetObjByName
; _LOCalc_RangeColumnGetObjByPosition
; _LOCalc_RangeColumnInsert
; _LOCalc_RangeColumnPageBreak
; _LOCalc_RangeColumnsGetCount
; _LOCalc_RangeColumnVisible
; _LOCalc_RangeColumnWidth
; _LOCalc_RangeCompute
; _LOCalc_RangeCopyMove
; _LOCalc_RangeCreateCursor
; _LOCalc_RangeData
; _LOCalc_RangeDelete
; _LOCalc_RangeFill
; _LOCalc_RangeFillSeries
; _LOCalc_RangeFilter
; _LOCalc_RangeFilterClear
; _LOCalc_RangeFindAll
; _LOCalc_RangeFindNext
; _LOCalc_RangeFormula
; _LOCalc_RangeGetAddressAsName
; _LOCalc_RangeGetAddressAsPosition
; _LOCalc_RangeGetCellByName
; _LOCalc_RangeGetCellByPosition
; _LOCalc_RangeGetSheet
; _LOCalc_RangeInsert
; _LOCalc_RangeNumbers
; _LOCalc_RangeQueryColumnDiff
; _LOCalc_RangeQueryContents
; _LOCalc_RangeQueryDependents
; _LOCalc_RangeQueryEmpty
; _LOCalc_RangeQueryFormula
; _LOCalc_RangeQueryIntersection
; _LOCalc_RangeQueryPrecedents
; _LOCalc_RangeQueryRowDiff
; _LOCalc_RangeQueryVisible
; _LOCalc_RangeReplace
; _LOCalc_RangeReplaceAll
; _LOCalc_RangeRowDelete
; _LOCalc_RangeRowGetObjByPosition
; _LOCalc_RangeRowHeight
; _LOCalc_RangeRowInsert
; _LOCalc_RangeRowPageBreak
; _LOCalc_RangeRowsGetCount
; _LOCalc_RangeRowVisible
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeClearContents
; Description ...: Clear specific cell contents in a range.
; Syntax ........: _LOCalc_RangeClearContents(ByRef $oRange, $iFlags)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell to clear the contents of. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iFlags              - an integer value (1-1023). The Cell Content type to delete. Can be BitOR'd together. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFlags not an Integer, less than 1, or greater than 1023. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Contents specified was successfully cleared from the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeClearContents(ByRef $oRange, $iFlags)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iFlags, $LOC_CELL_FLAG_VALUE, $LOC_CELL_FLAG_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRange.clearContents($iFlags)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeClearContents

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnDelete
; Description ...: Delete Columns from a Range.
; Syntax ........: _LOCalc_RangeColumnDelete(ByRef $oRange, $iColumn[, $iCount = 1])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iColumn             - an integer value. The column to begin deleting at. The Column called will be deleted. See remarks.
;                  $iCount              - [optional] an integer value. Default is 1. The number of columns to delete after the called column.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iColumns not an Integer or less than 0, or greater than number of Columns contained in the Range.
;				   @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested columns.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Columns in L.O. Calc are 0 based, to Delete Column "A" in the LibreOffice UI, you would call $iColumn with 0.
;				   Deleting Columns does not decrease the Column count, it simply erases the Column's contents in a specific area and shifts all after content left.
; Related .......: _LOCalc_RangeColumnInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnDelete(ByRef $oRange, $iColumn, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oRange.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iColumn, 0, $oColumns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iCount, 1, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumns.removeByIndex($iColumn, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeColumnDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnGetName
; Description ...: Retrieve the Column's name.
; Syntax ........: _LOCalc_RangeColumnGetName(ByRef $oColumn)
; Parameters ....: $oColumn             - [in/out] an object. A Column object returned by a previous _LOCalc_RangeColumnGetObjByPosition, or _LOCalc_RangeColumnGetObjByName function.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve the Column's name.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Success, returning Column's name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnGetName(ByRef $oColumn)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sName = $oColumn.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sName)
EndFunc   ;==>_LOCalc_RangeColumnGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnGetObjByName
; Description ...: Retrieve a Column's Object by name.
; Syntax ........: _LOCalc_RangeColumnGetObjByName(ByRef $oRange, $sName)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $sName               - a string value. The Column name to retrieve the Object for, such as "A".
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sName not a String.
;				   @Error 1 @Extended 3 Return 0 = Range does not contain a column with name called in $sName.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Column Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Success, returning Column's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeColumnGetObjByPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnGetObjByName(ByRef $oRange, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns, $oColumn

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumns = $oRange.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not $oColumns.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumn = $oColumns.getByName($sName)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOCalc_RangeColumnGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnGetObjByPosition
; Description ...: Retrieve the Column's Object by its position.
; Syntax ........: _LOCalc_RangeColumnGetObjByPosition(ByRef $oRange, $iColumn)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iColumn             - an integer value. The Column number to retrieve the Object for. See remarks.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, or less than 0, or greater than number of columns contained in the Range.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Column Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Success, returning Column's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Columns in L.O. Calc are 0 based, to retrieve Column "A" in the LibreOffice UI, you would call $iColumn with 0.
; Related .......: _LOCalc_RangeColumnGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnGetObjByPosition(ByRef $oRange, $iColumn)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns, $oColumn

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oRange.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iColumn, 0, $oColumns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumn = $oColumns.getByIndex($iColumn)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOCalc_RangeColumnGetObjByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnInsert
; Description ...: Insert blank columns into a Range at a specific column.
; Syntax ........: _LOCalc_RangeColumnInsert(ByRef $oRange, $iColumn[, $iCount = 1])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iColumn             - an integer value. The Column to insert the new column(s) at. See remarks. New columns will be inserted starting at this column and all content will be shifted right.
;                  $iCount              - [optional] an integer value. Default is 1. The number of blank columns to insert after the Column called.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iColumn not an Integer or less than 0, or greater than number of Columns contained in the Range.
;				   @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully inserted blank columns.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Columns in L.O. Calc are 0 based, to add columns in Column "A" in the LibreOffice UI, you would call $iColumn with 0.
;				   Inserting Columns does not increase the Column count, it simply adds blanks in a specific area and shifts all after content further right.
; Related .......: _LOCalc_RangeColumnDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnInsert(ByRef $oRange, $iColumn, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oRange.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iColumn, 0, $oColumns.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iCount, 1, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumns.insertByIndex($iColumn, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeColumnInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnPageBreak
; Description ...: Set or retrieve Page Break settings for a Column.
; Syntax ........: _LOCalc_RangeColumnPageBreak(ByRef $oColumn[, $bManualPageBreak = Null[, $bStartOfPageBreak = Null]])
; Parameters ....: $oColumn             - [in/out] an object. A Column object returned by a previous _LOCalc_RangeColumnGetObjByPosition, or _LOCalc_RangeColumnGetObjByName function.
;                  $bManualPageBreak    - [optional] a boolean value. Default is Null. If True, this column is the beginning of a manual Page Break.
;                  $bStartOfPageBreak   - [optional] a boolean value. Default is Null. If True, this column is the beginning of a start of Page Break. See Remarks.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bManualPageBreak not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bStartOfPageBreak not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $bManualPageBreak
;				   |								2 = Error setting $bStartOfPageBreak
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Setting $bStartOfPageBreak to True will insert a Manual Page Break, the same as setting $bManualPageBreak to True would.
;				   $bStartOfPageBreak setting is available more for indicating where Calc is inserting Page Breaks rather than for applying a setting. You can retrieve the settings for each Column, and check if this value is set to True or not. If the Page break is an automatically inserted one, the value for $bManualPageBreak would be false.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnPageBreak(ByRef $oColumn, $bManualPageBreak = Null, $bStartOfPageBreak = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abBreak[2]

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOCalc_VarsAreNull($bManualPageBreak, $bStartOfPageBreak) Then
		__LOCalc_ArrayFill($abBreak, $oColumn.IsManualPageBreak(), $oColumn.IsStartOfNewPage())

		Return SetError($__LO_STATUS_SUCCESS, 1, $abBreak)
	EndIf

	If ($bManualPageBreak <> Null) Then
		If Not IsBool($bManualPageBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oColumn.IsManualPageBreak = $bManualPageBreak
		$iError = ($oColumn.IsManualPageBreak() = $bManualPageBreak) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bStartOfPageBreak <> Null) Then
		If Not IsBool($bStartOfPageBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oColumn.IsStartOfNewPage = $bStartOfPageBreak
		$iError = ($oColumn.IsStartOfNewPage() = $bStartOfPageBreak) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeColumnPageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnsGetCount
; Description ...: Retrieve the total count of Columns contained in a Range.
; Syntax ........: _LOCalc_RangeColumnsGetCount(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Returning number of Columns contained in the Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: There is a fixed number of Columns per sheet, but different L.O. versions contain different amounts of Columns. But this also helps to determine how many columns are contained in a Cell Range.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnsGetCount(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oRange.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumns.Count())
EndFunc   ;==>_LOCalc_RangeColumnsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnVisible
; Description ...: Set or Retrieve the Column's visibility setting.
; Syntax ........: _LOCalc_RangeColumnVisible(ByRef $oColumn[, $bVisible = Null])
; Parameters ....: $oColumn             - an object. A Column object returned by a previous _LOCalc_RangeColumnGetObjByPosition, or _LOCalc_RangeColumnGetObjByName function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the Column is Visible.
; Return values .: Success: 1 or Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $bVisible
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Boolean = Success. All optional parameters were set to Null, returning Column's current visibility setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnVisible(ByRef $oColumn, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bVisible = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oColumn.IsVisible())


	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$oColumn.IsVisible = $bVisible
	$iError = ($oColumn.IsVisible() = $bVisible) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeColumnVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeColumnWidth
; Description ...: Set or Retrieve the Column's Width settings.
; Syntax ........: _LOCalc_RangeColumnWidth(ByRef $oColumn[, $bOptimal = Null[, $iWidth = Null]])
; Parameters ....: $oColumn             - an object. A Column object returned by a previous _LOCalc_RangeColumnGetObjByPosition, or _LOCalc_RangeColumnGetObjByName function.
;                  $bOptimal            - [optional] a boolean value. Default is Null. If True, the Optimal width is automatically chosen. See Remarks.
;                  $iWidth              - [optional] an integer value (0-34464). Default is Null. The Width of the Column, set in Micrometers.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bOptimal not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $iWidth not an Integer, less than 0 or greater than 34464.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $bOptimal
;				   |								2 = Error setting $iWidth
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bOptimal only accepts True. False will return an error. Calling True again returns the cell to optimal width, setting a custom width essentially disables it.
;				   Note: I am presently unable to find a setting for Optimal Width "Add" Value.
; Related .......: _LOCalc_ConvertFromMicrometer, _LOCalc_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeColumnWidth(ByRef $oColumn, $bOptimal = Null, $iWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avWidth[2]
	Local $iError = 0

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOCalc_VarsAreNull($bOptimal, $iWidth) Then
		__LOCalc_ArrayFill($avWidth, $oColumn.OptimalWidth(), $oColumn.Width())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avWidth)
	EndIf

	If ($bOptimal <> Null) Then
		If Not IsBool($bOptimal) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oColumn.OptimalWidth = $bOptimal
		$iError = ($oColumn.OptimalWidth() = $bOptimal) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iWidth <> Null) Then
		If Not __LOCalc_IntIsBetween($iWidth, 0, 34464) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oColumn.Width = $iWidth
		$iError = (__LOCalc_IntIsBetween($oColumn.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeColumnWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeCompute
; Description ...: Perform a Computation function on a Range. See Remarks.
; Syntax ........: _LOCalc_RangeCompute(ByRef $oRange, $iFunction)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iFunction           - an integer value (0-12). The Computation Function to perform. See Constants $LOC_COMPUTE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Number
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFunction not an Integer, less than 0 or greater than 12. See Constants $LOC_COMPUTE_* as defined in LibreOfficeCalc_Constants.au3.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to perform computation.
;				   --Success--
;				   @Error 0 @Extended 0 Return Number = Success. Successfully performed the requested computation, returning the result as a Numerical value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This makes no changes in the document itself, it only returns the result of the computation.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeCompute(ByRef $oRange, $iFunction)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $nResult

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iFunction, $LOC_COMPUTE_NONE, $LOC_COMPUTE_VARP) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$nResult = $oRange.computeFunction($iFunction)
	If Not IsNumber($nResult) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $nResult)
EndFunc   ;==>_LOCalc_RangeCompute

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeCopyMove
; Description ...: Copy or Move a Cell or Cell Range to another range.
; Syntax ........: _LOCalc_RangeCopyMove(ByRef $oSheet, ByRef $oRangeSrc, ByRef $oRangeDest[, $bMove = False])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oRangeSrc           - [in/out] an object. The Cell or Cell Range to copy or move from. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oRangeDest          - [in/out] an object. The Cell or Cell Range to copy or move to. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bMove               - [optional] a boolean value. Default is False. If True, the cell range is moved to the destination. If False, the Cell Range is only copied.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oRangeSrc not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oRangeDest not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bMove not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Source Cell Range Address.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Destination Cell Range Address.
;				   @Error 2 @Extended 3 Return 0 = Failed to create "com.sun.star.table.CellAddress" Struct.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Cell or Cell range was successfully copied to destination.
;				   @Error 0 @Extended 1 Return 1 = Success. Cell or Cell range was successfully moved to destination.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Destination Range can be on different Sheet from Source.
;				   $oSheet is the source sheet where $oRangeSrc is located.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeCopyMove(ByRef $oSheet, ByRef $oRangeSrc, ByRef $oRangeDest, $bMove = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tCellSrcAddr, $tInputCellDestAddr, $tCellDestAddr

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRangeSrc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oRangeDest) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bMove) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$tCellSrcAddr = $oRangeSrc.RangeAddress()
	If Not IsObj($tCellSrcAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tInputCellDestAddr = $oRangeDest.RangeAddress()
	If Not IsObj($tInputCellDestAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$tCellDestAddr = __LOCalc_CreateStruct("com.sun.star.table.CellAddress")
	If Not IsObj($tCellDestAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	$tCellDestAddr.Sheet = $tInputCellDestAddr.Sheet()
	$tCellDestAddr.Column = $tInputCellDestAddr.StartColumn()
	$tCellDestAddr.Row = $tInputCellDestAddr.StartRow()

	If $bMove Then
		$oSheet.MoveRange($tCellDestAddr, $tCellSrcAddr)
		Return SetError($__LO_STATUS_SUCCESS, 1, 1)

	Else
		$oSheet.CopyRange($tCellDestAddr, $tCellSrcAddr)
		Return SetError($__LO_STATUS_SUCCESS, 0, 1)

	EndIf
EndFunc   ;==>_LOCalc_RangeCopyMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeCreateCursor
; Description ...: Create a Sheet Cursor for a particular range.
; Syntax ........: _LOCalc_RangeCreateCursor(ByRef $oSheet, ByRef $oRange)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create a Sheet Cursor.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully created a Sheet Cursor for the specified Range, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Sheet Cursor can be used in functions accepting a range. When created, the Cursor will have the called range selected.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeCreateCursor(ByRef $oSheet, ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheetCursor

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheetCursor = $oSheet.createCursorByRange($oRange)
	If Not IsObj($oSheetCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheetCursor)
EndFunc   ;==>_LOCalc_RangeCreateCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeData
; Description ...: Set or Retrieve Data in a Range.
; Syntax ........: _LOCalc_RangeData(ByRef $oRange[, $aavData = Null])
; Parameters ....: $oRange              - [in/out] an object. The Cell or Cell Range to set or retrieve data . A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $aavData             - [optional] an array of Arrays containing variants. Default is Null. An Array of Arrays containing data, strings or numbers, to fill the range with. See remarks.
; Return values .: Success: 1 or Array
;				   Failure: 0 or ? and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $aavData not an Array.
;				   @Error 1 @Extended 3 Return 0 = $aavData array contains less or more elements than number of rows contained in the cell range.
;				   @Error 1 @Extended 4 Return ? = Element of $aavData does not contain an array. Returning array element number of $aavData containing error.
;				   @Error 1 @Extended 5 Return ? = Array contained in $aavData has less or more elements than number of columns in the cell range. Returning array element number of $aavData containing faulty array.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve array of Data contained in the Cell Range.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Start of Row from Cell Range.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve End of Row from Cell Range.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve Start of Column from Cell Range.
;				   @Error 2 @Extended 5 Return 0 = Failed to retrieve End of Column from Cell Range.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Data was successfully set for the cell range.
;				   @Error 0 @Extended 1 Return Array of Arrays = Success. $aavData set to Null, returning an array containing arrays, which contain any data content contained in the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will return Strings and Numbers contained in the cell range when $aavData is called with Null keyword. Array will be an array of arrays. The internal arrays will contain numerical or string data, depending on cell content.
;				   $aavData must be an array containing arrays. The main Array's element count must match the row count contained in the Cell Range, and each internal Array's element count must match the column count of the Cell Range it is to fill.
;				   Any data previously contained in the Cell Range will be overwritten.
;				   All array elements must contain appropriate data, strings or numbers.
;				   Formulas will be inserted as strings only, and will not be valid.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeData(ByRef $oRange, $aavData = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iStart, $iEnd

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($aavData = Null) Then
		$aavData = $oRange.getDataArray()
		If Not IsArray($aavData) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		Return SetError($__LO_STATUS_SUCCESS, 1, $aavData)
	EndIf

	If Not IsArray($aavData) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	; Determine if the Array is sized appropriately
	$iStart = $oRange.RangeAddress.StartRow()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$iEnd = $oRange.RangeAddress.EndRow()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If (UBound($aavData) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iStart = $oRange.RangeAddress.StartColumn()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$iEnd = $oRange.RangeAddress.EndColumn()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)

	For $i = 0 To UBound($aavData) - 1
		If Not IsArray($aavData[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
		If (UBound($aavData[$i]) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
	Next

	$oRange.setDataArray($aavData)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeDelete
; Description ...: Delete a Range of cell contents and reposition surrounding cells.
; Syntax ........: _LOCalc_RangeDelete(ByRef $oSheet, $oRange, $iMode)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oRange              - an object. A Cell or Cell range to delete. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iMode               - an integer value (0-4). The Cell Deletion Mode. See Constants $LOC_CELL_DELETE_MODE_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iMode not an Integer, less than 0, or greater than 4. See Constants $LOC_CELL_DELETE_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Range Address Struct from $oRange.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Cell range was successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will silently fail if the deletion will cause an array formula to be split -- OOME. 4.1., Page 509.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeDelete(ByRef $oSheet, $oRange, $iMode)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tCellAddr

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iMode, $LOC_CELL_DELETE_MODE_NONE, $LOC_CELL_DELETE_MODE_COLUMNS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddr = $oRange.RangeAddress()
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oSheet.removeRange($tCellAddr, $iMode)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFill
; Description ...: Automatically fill cells with a value. See Remarks.
; Syntax ........: _LOCalc_RangeFill(ByRef $oRange, $iDirection[, $iCount = 1])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iDirection          - an integer value (0-3). The Direction to perform the Fill operation. See Constants $LOC_FILL_DIR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iCount              - [optional] an integer value. Default is 1. The number of Cells to take into account at the beginning of the range to constitute the fill algorithm.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iDirection not an Integer, less than 0 or greater than 3. See Constants $LOC_FILL_DIR_* as defined in LibreOfficeCalc_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Fill operation was successfully processed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Fill value is calculated based on the first value(s) in the Range, the first value location depends on the Fill direction. If Fill direction is set to Right, the initial value must be in the first cell(s) on the left, and vice versa.
; Related .......: _LOCalc_RangeFillSeries
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFill(ByRef $oRange, $iDirection, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iDirection, $LOC_FILL_DIR_DOWN, $LOC_FILL_DIR_LEFT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iCount, 0, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oRange.fillAuto($iDirection, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFill

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFillSeries
; Description ...: Fill a Range of Cells with Data.
; Syntax ........: _LOCalc_RangeFillSeries(ByRef $oRange, $iDirection, $iMode, $nStep, $nEnd[, $iDateMode = $LOC_FILL_DATE_MODE_DAY])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iDirection          - an integer value (0-3). The Direction of the Series Fill. See Constants $LOC_FILL_DIR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iMode               - an integer value (0-4). The Fill Type. See Constants $LOC_FILL_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $nStep               - a general number value. The amount the beginning value increments per step.
;                  $nEnd                - a general number value. The maximum Value the Fill series can insert.
;                  $iDateMode           - [optional] an integer value (0-3). Default is $LOC_FILL_DATE_MODE_DAY. The mode to calculate dates if $iMode is set to $LOC_FILL_MODE_DATE. See Constants $LOC_FILL_DATE_MODE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iDirection not an Integer, less than 0 or greater than 3. See Constants $LOC_FILL_DIR_* as defined in LibreOfficeCalc_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iMode not an Integer, less than 0 or greater than 4. See Constants $LOC_FILL_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $nStep not a Number value.
;				   @Error 1 @Extended 5 Return 0 = $nEnd not a Number value.
;				   @Error 1 @Extended 6 Return 0 = $iDateMode not an Integer, less than 0 or greater than 3. See Constants $LOC_FILL_DATE_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Fill series was successfully processed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeFill
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFillSeries(ByRef $oRange, $iDirection, $iMode, $nStep, $nEnd, $iDateMode = $LOC_FILL_DATE_MODE_DAY)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iDirection, $LOC_FILL_DIR_DOWN, $LOC_FILL_DIR_LEFT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iMode, $LOC_FILL_MODE_SIMPLE, $LOC_FILL_MODE_AUTO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsNumber($nStep) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsNumber($nEnd) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not __LOCalc_IntIsBetween($iDateMode, $LOC_FILL_DATE_MODE_DAY, $LOC_FILL_DATE_MODE_YEAR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oRange.fillSeries($iDirection, $iMode, $iDateMode, $nStep, $nEnd)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFillSeries

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFilter
; Description ...: Apply a Filter to a Range.
; Syntax ........: _LOCalc_RangeFilter(ByRef $oRange, ByRef $oFilterDesc)
; Parameters ....: $oRange              - [in/out] an object. The Range to Filter. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oFilterDesc         - [in/out] an object. A Filter Descriptor created by a previous _LOCalc_FilterDescriptorCreate function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oFilterDesc not an Object.
;				   @Error 1 @Extended 3 Return 0 = Object called in $oFilterDesc not a Filter Descriptor.
;				   @Error 1 @Extended 4 Return ? = Column called in one Filter Field is greater than number of columns in the Range. Returning FilterFields Array element containing bad Filter Field, as an Integer.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve Filter Fields array from Filter Descriptor.
;				   @Error 3 @Extended 2 Return 0 = Failed to get count of columns contained in Range.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully processed Filter operation.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeFilterClear, _LOCalc_FilterDescriptorCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFilter(ByRef $oRange, ByRef $oFilterDesc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iColumns
	Local $atFilterField[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFilterDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oFilterDesc.supportsService("com.sun.star.sheet.SheetFilterDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$atFilterField = $oFilterDesc.getFilterFields2()
	If Not IsArray($atFilterField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iColumns = $oRange.Columns.Count()
	If Not IsInt($iColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($atFilterField) - 1
		If Not __LOCalc_IntIsBetween($atFilterField[$i].Field(), 0, $iColumns - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
	Next

	$oRange.filter($oFilterDesc)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFilter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFilterClear
; Description ...: Clear any previous filters for a Range.
; Syntax ........: _LOCalc_RangeFilterClear(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. The Range to clear filtering for. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create a new, blank, Filter Descriptor.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully cleared any old Filters for the Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_RangeFilter
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFilterClear(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFilterDesc

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oFilterDesc = $oRange.createFilterDescriptor(True)
	If Not IsObj($oFilterDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oRange.filter($oFilterDesc) ; Calling filter with a blank Filter Desc clears any old filters applied.

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFilterClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFindAll
; Description ...: Find all matches contained in a document of a specified Search String.
; Syntax ........: _LOCalc_RangeFindAll(ByRef $oRange, ByRef $oSrchDescript, $sSearchString)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or regular expression to search for.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSrchDescriptObject not a Search Descriptor Object.
;				   @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Search did not return an Object, something went wrong.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Search was Successful, but found no results.
;				   @Error 0 @Extended ? Return Array = Success. Search was Successful, returning 1 dimensional array containing the objects to each match, @Exteneded is set to the number of matches.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Objects returned are Ranges and can be used in any of the functions accepting a Range Object etc., to modify their properties or even the text itself.
;				   Only the Sheet that contains the Range is searched, to search all Sheets you will have to cycle through and perform a search for each.
; Related .......: _LOCalc_SearchDescriptorCreate, _LOCalc_RangeFindNext, _LOCalc_RangeReplaceAll, _LOCalc_RangeReplace
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFindAll(ByRef $oRange, ByRef $oSrchDescript, $sSearchString)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResults
	Local $aoResults[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oSrchDescript.SearchString = $sSearchString

	$oResults = $oRange.findAll($oSrchDescript)
	If Not IsObj($oResults) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oResults.getCount() > 0) Then
		ReDim $aoResults[$oResults.getCount]
		For $i = 0 To $oResults.getCount() - 1
			$aoResults[$i] = $oResults.getByIndex($i)
			Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	Return (UBound($aoResults) > 0) ? (SetError($__LO_STATUS_SUCCESS, UBound($aoResults), $aoResults)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeFindAll

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFindNext
; Description ...: Find a Search String in a Document once or one at a time.
; Syntax ........: _LOCalc_RangeFindNext(ByRef $oRange, ByRef $oSrchDescript, $sSearchString[, $oLastFind = Null])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $oLastFind           - [optional] an object. Default is Null. The last returned Object by a previous call to this function to begin the search from, if set to Null, the search begins at the start of the Range.
; Return values .: Success: Object or 1.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;				   @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;				   @Error 1 @Extended 5 Return 0 = $oLastFind not an Object and not set to Null, or failed to retrieve starting position from $oRange.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Search was successful but found no matches.
;				   @Error 0 @Extended 1 Return Object = Success. Search was successful, returning the resulting Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Object returned is a Range and can be used in any of the functions accepting a Range Object etc., to modify their properties or even the text itself.
;				   Only the Sheet that contains the Range is searched, to search all Sheets you will have to cycle through and perform a search for each.
; Related .......: _LOCalc_SearchDescriptorCreate, _LOCalc_RangeFindAll, _LOCalc_RangeReplaceAll, _LOCalc_RangeReplace
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFindNext(ByRef $oRange, ByRef $oSrchDescript, $sSearchString, $oLastFind = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResult, $oFindRange

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If ($oLastFind <> Null) Then ; If Last find is set, set search start for beginning or end of last result, depending SearchBackwards value.
		If Not IsObj($oLastFind) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		; If Search Backwards is False, then retrieve the end of the last result's range, else get the Start.
		$oFindRange = ($oSrchDescript.SearchBackwards() = False) ? ($oRange.getCellByPosition($oLastFind.RangeAddress.EndColumn(), $oLastFind.RangeAddress.EndRow())) : ($oRange.getCellByPosition($oLastFind.RangeAddress.StartColumn(), $oLastFind.RangeAddress.StartRow()))
	EndIf

	$oSrchDescript.SearchString = $sSearchString

	If IsObj($oLastFind) Then
		$oResult = $oRange.findNext($oFindRange, $oSrchDescript)

	Else ; If a search hasn't been done before, FindFirst must be used or a result could be missed in the first cell.
		$oResult = $oRange.findFirst($oSrchDescript)
	EndIf

	Return (IsObj($oResult)) ? (SetError($__LO_STATUS_SUCCESS, 1, $oResult)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeFindNext

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeFormula
; Description ...: Set or Retrieve Formulas in a Range.
; Syntax ........: _LOCalc_RangeFormula(ByRef $oRange[, $aasFormulas = Null])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $aasFormulas         - [optional] an array or arrays containing strings. Default is Null. An Array of Arrays containing formula strings to fill the range with. See remarks.
; Return values .: Success: 1 or Array
;				   Failure: 0 or ? and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $aasFormulas not an Array.
;				   @Error 1 @Extended 3 Return 0 = $aasFormulas array contains less or more elements than number of rows contained in the cell range.
;				   @Error 1 @Extended 4 Return ? = Element of $aasFormulas does not contain an array. Returning array element number of $aasFormulas containing error.
;				   @Error 1 @Extended 5 Return ? = Array contained in $aasFormulas has less or more elements than number of columns in the cell range. Returning array element number of $aasFormulas containing faulty array.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve array of Formula Data contained in the Cell Range.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Start of Row from Cell Range.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve End of Row from Cell Range.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve Start of Column from Cell Range.
;				   @Error 2 @Extended 5 Return 0 = Failed to retrieve End of Column from Cell Range.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Formulas were successfully set for the cell range.
;				   @Error 0 @Extended 1 Return Array of Arrays = Success. $aasFormulas set to Null, returning an array containing arrays, which contain any Formula content contained in the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will return only formulas contained in the cell range when $aasFormulas is called with Null keyword. Array will be an array of arrays. The internal arrays will contain blank cells or formula strings, depending on cell content.
;				   $aasFormulas must be an array containing arrays. The main Array's element count must match the row count contained in the Cell Range, and each internal Array's element count must match the column count of the Cell Range it is to fill.
;				   Any data previously contained in the Cell Range will be overwritten.
;				   All array elements must contain strings, blank or otherwise.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeFormula(ByRef $oRange, $aasFormulas = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iStart, $iEnd

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($aasFormulas = Null) Then
		$aasFormulas = $oRange.getFormulaArray()
		If Not IsArray($aasFormulas) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		Return SetError($__LO_STATUS_SUCCESS, 1, $aasFormulas)
	EndIf

	If Not IsArray($aasFormulas) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	; Determine if the Array is sized appropriately
	$iStart = $oRange.RangeAddress.StartRow()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$iEnd = $oRange.RangeAddress.EndRow()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If (UBound($aasFormulas) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iStart = $oRange.RangeAddress.StartColumn()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$iEnd = $oRange.RangeAddress.EndColumn()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)

	For $i = 0 To UBound($aasFormulas) - 1
		If Not IsArray($aasFormulas[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
		If (UBound($aasFormulas[$i]) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
	Next

	$oRange.setFormulaArray($aasFormulas)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGetAddressAsName
; Description ...: Retrieve the Name of the beginning and ending cells of the range.
; Syntax ........: _LOCalc_RangeGetAddressAsName(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to Retrieve Range Address.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Successfully retrieved Range's address, returning it as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Return will be like the following, including the dollar signs. "$Sheet1.$A$1:$F$18"
; Related .......: _LOCalc_RangeGetAddressAsPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGetAddressAsName(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sName = $oRange.AbsoluteName()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sName)
EndFunc   ;==>_LOCalc_RangeGetAddressAsName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGetAddressAsPosition
; Description ...: Retrieve the Position of the beginning and ending cells of the range.
; Syntax ........: _LOCalc_RangeGetAddressAsPosition(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to Retrieve Range Address.
;				   --Success--
;				   @Error 0 @Extended 0 Return Array = Success. Successfully retrieved Range's address, returning it as a 5 element Array. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The return will be a 5 element array giving the Range's address in the following order: Sheet index number, Range's first Cell Column, First Cell Row, Last Cell Column, Last Cell Row.
; Related .......: _LOCalc_RangeGetAddressAsName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGetAddressAsPosition(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tRangeAddr
	Local $aiAddress[5]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tRangeAddr = $oRange.RangeAddress()
	If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aiAddress[0] = $tRangeAddr.Sheet()
	$aiAddress[1] = $tRangeAddr.StartColumn()
	$aiAddress[2] = $tRangeAddr.StartRow()
	$aiAddress[3] = $tRangeAddr.EndColumn()
	$aiAddress[4] = $tRangeAddr.EndRow()

	Return SetError($__LO_STATUS_SUCCESS, 0, $aiAddress)
EndFunc   ;==>_LOCalc_RangeGetAddressAsPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGetCellByName
; Description ...: Retrieve a Cell or Cell Range Object by Cell name.
; Syntax ........: _LOCalc_RangeGetCellByName(ByRef $oRange, $sFromCellName[, $sToCellName = Null])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $sFromCellName       - a string value. The cell to retrieve the Object for, or to begin the Cell Range. See remarks.
;                  $sToCellName         - [optional] a string value. Default is Null. The cell to end the Cell Range at.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFromCellName not a String.
;				   @Error 1 @Extended 3 Return 0 = $sToCellName not set to Null, and not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve requested Cell or Cell Range Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully retrieved and returning requested Cell or Cell Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFromCellName can be a Cell Name or a defined Cell Range name.
; Related .......: _LOCalc_RangeGetCellByPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGetCellByName(ByRef $oRange, $sFromCellName, $sToCellName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCellRange
	Local $sCellRange = $sFromCellName

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFromCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($sToCellName <> Null) And Not IsString($sToCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sToCellName <> Null) Then $sCellRange &= ":" & $sToCellName

	$oCellRange = $oRange.getCellRangeByName($sCellRange)
	If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCellRange)
EndFunc   ;==>_LOCalc_RangeGetCellByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGetCellByPosition
; Description ...: Retrieve a Cell or Cell Range Object by position.
; Syntax ........: _LOCalc_RangeGetCellByPosition(ByRef $oRange, $iColumn, $iRow[, $iToColumn = Null[, $iToRow = Null]])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iColumn             - an integer value. The Column of the desired cell, or of the beginning of the Cell range. 0 Based. See remarks.
;                  $iRow                - an integer value. The Row of the desired cell, or of the beginning of the Cell range. 0 Based. See remarks.
;                  $iToColumn           - [optional] an integer value. Default is Null. The Column of the end of the Cell range. 0 Based. Must be greater or equal to $iColumn.
;                  $iToRow              - [optional] an integer value. Default is Null. The Row of the end of the Cell range. 0 Based. Must be greater or equal to $iRow.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, or less than 0, or greater than number of Columns contained in the Range.
;				   @Error 1 @Extended 3 Return 0 = $iRow not an Integer, or less than 0, or greater than number of Rows contained in the Range.
;				   @Error 1 @Extended 4 Return 0 = $iToColumn not an Integer, or less than 0, or greater than number of Columns contained in the Range.
;				   @Error 1 @Extended 5 Return 0 = $iToRow not an Integer, or less than 0, or greater than number of Rows contained in the Range.
;				   @Error 1 @Extended 6 Return 0 = $iToColumn less than $iColumn.
;				   @Error 1 @Extended 7 Return 0 = $iToRow less than $iRow.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve an individual Cell's Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve a Cell Range's Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully retrieved and returned an Individual Cell's Object.
;				   @Error 0 @Extended 1 Return Object = Success. Successfully retrieved and returned a Cell Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: According to the wiki (https://wiki.documentfoundation.org/Faq/Calc/022), the maximum Columns contained in a sheet is 1024 until version 7.3, or 16384 from 7.3. and up..
;				   According to Andrew Pitonyak, (OOME. 4.1 Page 492), the maximum number of rows contained in a sheet is 65,536 as of OOo Calc 3.0, but according to the wiki (https://wiki.documentfoundation.org/Faq/Calc/022), the maximum or Rows for Libre Office Calc is 1,048,576.
; Related .......: _LOCalc_RangeGetCellByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGetCellByPosition(ByRef $oRange, $iColumn, $iRow, $iToColumn = Null, $iToRow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell, $oCellRange

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iColumn, 0, $oRange.Columns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iRow, 0, $oRange.Rows.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($iToColumn <> Null) Or ($iToRow <> Null) Then
		If Not __LOCalc_IntIsBetween($iToColumn, 0, $oRange.Columns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If Not __LOCalc_IntIsBetween($iToRow, 0, $oRange.Rows.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If ($iToColumn < $iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If ($iToRow < $iRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	EndIf

	If ($iToColumn = Null) And ($iToRow = Null) Then
		$oCell = $oRange.getCellByPosition($iColumn, $iRow)
		If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		Return SetError($__LO_STATUS_SUCCESS, 0, $oCell)

	Else
		$oCellRange = $oRange.getCellRangeByPosition($iColumn, $iRow, $iToColumn, $iToRow)
		If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
		Return SetError($__LO_STATUS_SUCCESS, 1, $oCellRange)

	EndIf
EndFunc   ;==>_LOCalc_RangeGetCellByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeGetSheet
; Description ...: Return the Sheet Object that contains the Range.
; Syntax ........: _LOCalc_RangeGetSheet(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to Retrieve Sheet Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully retrieved Range's parent Sheet, returning the Sheet's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeGetSheet(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheet = $oRange.Spreadsheet()
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheet)
EndFunc   ;==>_LOCalc_RangeGetSheet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeInsert
; Description ...: Insert blank cells at a Cell Range.
; Syntax ........: _LOCalc_RangeInsert(ByRef $oSheet, $oRange, $iMode)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_SheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oRange              - an object. A Cell or Cell Range to insert new blank cells at. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iMode               - an integer value (0-4). The Cell Insertion Mode. See Constants $LOC_CELL_INSERT_MODE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iMode not an Integer, less than 0, or greater than 4. See Constants $LOC_CELL_INSERT_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Range Address Struct from $oRange.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Blank cells were successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: The new range of cells inserted will be the same size as the range called in $oRange.
;				   Non-Empty cells cannot be moved off of the sheet.
;				   This function will silently fail if the insertion will cause an array formula to be split -- OOME. 4.1., Page 509.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeInsert(ByRef $oSheet, $oRange, $iMode)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tCellAddr

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iMode, $LOC_CELL_INSERT_MODE_NONE, $LOC_CELL_INSERT_MODE_COLUMNS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddr = $oRange.RangeAddress()
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oSheet.insertCells($tCellAddr, $iMode)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeNumbers
; Description ...: Set or Retrieve Numbers in a Range.
; Syntax ........: _LOCalc_RangeNumbers(ByRef $oRange[, $aanNumbers = Null])
; Parameters ....: $oRange              - [in/out] an object. A cell or cell range to set or retrieve number values for. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $aanNumbers          - [optional] an array of arrays containing general numbers. Default is Null. An Array of Arrays containing numbers to fill the range with. See remarks.
; Return values .: Success: 1 or Array
;				   Failure: 0 or ? and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $aanNumbers not an Array.
;				   @Error 1 @Extended 3 Return 0 = $aanNumbers array contains less or more elements than number of rows contained in the cell range.
;				   @Error 1 @Extended 4 Return ? = Element of $aanNumbers does not contain an array. Returning array element number of $aanNumbers containing error.
;				   @Error 1 @Extended 5 Return ? = Array contained in $aanNumbers has less or more elements than number of columns in the cell range. Returning array element number of $aanNumbers containing faulty array.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve array of Numerical Data contained in the Cell Range.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Start of Row from Cell Range.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve End of Row from Cell Range.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve Start of Column from Cell Range.
;				   @Error 2 @Extended 5 Return 0 = Failed to retrieve End of Column from Cell Range.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Values were successfully set for the cell range.
;				   @Error 0 @Extended 1 Return Array of Arrays = Success. $aanNumbers set to Null, returning an array containing arrays, which contain any numerical content contained in the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will return only numbers contained in the cell range when $aanNumbers is called with Null keyword. Array will be an array of arrays. The internal arrays will contain blank cells or numbers, depending on cell content.
;				   $aanNumbers must be an array containing arrays. The main Array's element count must match the row count contained in the Cell Range, and each internal Array's element count must match the column count of the Cell Range it is to fill.
;				   Any data previously contained in the Cell Range will be overwritten.
;				   All array elements must contain numbers.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeNumbers(ByRef $oRange, $aanNumbers = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iStart, $iEnd

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($aanNumbers = Null) Then
		$aanNumbers = $oRange.getData()
		If Not IsArray($aanNumbers) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		Return SetError($__LO_STATUS_SUCCESS, 1, $aanNumbers)
	EndIf

	If Not IsArray($aanNumbers) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	; Determine if the Array is sized appropriately
	$iStart = $oRange.RangeAddress.StartRow()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$iEnd = $oRange.RangeAddress.EndRow()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If (UBound($aanNumbers) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iStart = $oRange.RangeAddress.StartColumn()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$iEnd = $oRange.RangeAddress.EndColumn()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)

	For $i = 0 To UBound($aanNumbers) - 1
		If Not IsArray($aanNumbers[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
		If (UBound($aanNumbers[$i]) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
	Next

	$oRange.setData($aanNumbers)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeNumbers

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryColumnDiff
; Description ...: Query a Cell Range for differences on each column based on a specific row.
; Syntax ........: _LOCalc_RangeQueryColumnDiff(ByRef $oRange, $oCellToCompare)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range to look for differences in. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oCellToCompare      - an object. A single Cell object (not a range) returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function. The Row this cell is located in will be used for the query.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCellToCompare not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCellToCompare is not a single cell, cell ranges are not supported.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Cell Address Struct from $oCellToCompare.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query column differences.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Looks for differences per column in the range, comparing the column to the value in the row $oCellToCompare is located. OOME 4.1. pg 488/489
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryColumnDiff(ByRef $oRange, $oCellToCompare)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $tCellAddr
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCellToCompare) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not ($oCellToCompare.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddr = $oCellToCompare.CellAddress()
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oRanges = $oRange.queryColumnDifferences($tCellAddr)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryColumnDiff

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryContents
; Description ...: Query a Cell or Cell range for specific cell contents.
; Syntax ........: _LOCalc_RangeQueryContents(ByRef $oRange, $iFlags)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iFlags              - an integer value (1-1023). The Cell content type flag. Can be BitOR'd together. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFlags not an Integer, less than 1 or greater than 1023. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query cell content.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Empty cells in the range may be skipped depending on the flag used. For instance, when querying for styles, the returned ranges may not include empty cells even if styles are applied to those cells.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryContents(ByRef $oRange, $iFlags)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iFlags, $LOC_CELL_FLAG_VALUE, $LOC_CELL_FLAG_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRanges = $oRange.queryContentCells($iFlags)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryContents

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryDependents
; Description ...: Query a Cell or Cell Range for Dependents.
; Syntax ........: _LOCalc_RangeQueryDependents(ByRef $oRange[, $bRecursive = False])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bRecursive          - [optional] a boolean value. Default is False. If True, the query is repeated for each found cell.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bRecursive not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query cell dependents.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Dependent cells are cells which reference cells in this range. If $bRecursive is True, repeats query with all found cells (finds dependents of dependents, and so on).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryDependents(ByRef $oRange, $bRecursive = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bRecursive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRanges = $oRange.queryDependents($bRecursive)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryDependents

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryEmpty
; Description ...: Query a Cell or Cell Range for empty cells.
; Syntax ........: _LOCalc_RangeQueryEmpty(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query for empty cells.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryEmpty(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRanges = $oRange.queryEmptyCells()
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryEmpty

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryFormula
; Description ...: Query a Cell or Cell Range for formulas having a specific result.
; Syntax ........: _LOCalc_RangeQueryFormula(ByRef $oRange, $iResultType)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iResultType         - an integer value (1-7). The Formula result type. Can be BitOR'd together. See Constants $LOC_FORMULA_RESULT_TYPE_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iResultType not an Integer, less than 1, or greater than 7. See Constants $LOC_FORMULA_RESULT_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query cell formula results.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryFormula(ByRef $oRange, $iResultType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iResultType, $LOC_FORMULA_RESULT_TYPE_VALUE, $LOC_FORMULA_RESULT_TYPE_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRanges = $oRange.queryFormulaCells($iResultType)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryIntersection
; Description ...: Retrieve an array of cell ranges that intersect with a certain cell range.
; Syntax ........: _LOCalc_RangeQueryIntersection(ByRef $oRange, $oCell)
; Parameters ....: $oRange              - [in/out] an object. A Cell range that contains the cell or cell range called in $oCell. A Cell Range object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oCell               - an object. A Cell or Cell Range located inside of the cell range called in $oRange. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Range Address Struct from $oCell.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query cell range intersections.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryIntersection(ByRef $oRange, $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $tRangeAddr
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tRangeAddr = $oCell.RangeAddress()
	If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oRanges = $oRange.queryIntersection($tRangeAddr)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oCell.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryIntersection

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryPrecedents
; Description ...: Query a Cell or Cell Range for Precedents.
; Syntax ........: _LOCalc_RangeQueryPrecedents(ByRef $oRange[, $bRecursive = False])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bRecursive          - [optional] a boolean value. Default is False. If True, the query is repeated for each found cell.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bRecursive not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query cell precedents.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Precedent cells are cells which are referenced by cells in this range. If $bRecursive is True, repeats query with all found cells (finds precedents of precedents, and so on).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryPrecedents(ByRef $oRange, $bRecursive = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bRecursive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRanges = $oRange.queryPrecedents($bRecursive)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryPrecedents

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryRowDiff
; Description ...: Query a Cell Range for differences on each row based on a specific column.
; Syntax ........: _LOCalc_RangeQueryRowDiff(ByRef $oRange, $oCellToCompare)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range to look for differences in. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oCellToCompare      - an object. A single Cell object (not a range) returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function. The Column this cell is located in will be used for the query.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCellToCompare not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCellToCompare is not a single cell, cell ranges are not supported.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Cell Address Struct from $oCellToCompare.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query row differences.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Looks for differences per row in the range, comparing the row to the value in the column $oCellToCompare is located. OOME 4.1. pg 488/489
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryRowDiff(ByRef $oRange, $oCellToCompare)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $tCellAddr
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCellToCompare) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not ($oCellToCompare.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddr = $oCellToCompare.CellAddress()
	If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oRanges = $oRange.queryRowDifferences($tCellAddr)
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryRowDiff

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeQueryVisible
; Description ...: Query a Cell or Cell Range for visible cells.
; Syntax ........: _LOCalc_RangeQueryVisible(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query for visible cell.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeQueryVisible(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRanges[0]

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRanges = $oRange.queryVisibleCells()
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aoRanges = $oRanges.getRangeAddresses()
	If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aoRanges) - 1
		$aoRanges[$i] = $oRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
		If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc   ;==>_LOCalc_RangeQueryVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeReplace
; Description ...: Replace the first instances of a search within a Range.
; Syntax ........: _LOCalc_RangeReplace(ByRef $oRange, ByRef $oSrchDescript, $sSearchString, $sReplaceString)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $sReplaceString      - a string value. A String of text or a regular expression to replace the first result with.
; Return values .: Success: 0 or Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;				   @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;				   @Error 1 @Extended 5 Return 0 = $sReplaceString not a String.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Found a result, but failed to replace it.
;				   --Success--
;				   @Error 0 @Extended 1 Return Object = Success. Search and Replace was successful, returning Object for Cell that the find and replace was performed upon.
;				   @Error 0 @Extended 0 Return 0 = Success. Search and replace was successful, no results found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Libre Office does not offer a Function to call to replace only one result within a Range, consequently I have had to create my own, which means this may not work exactly as expected.
; Related .......: _LOCalc_SearchDescriptorCreate, _LOCalc_RangeFindAll, _LOCalc_RangeFindNext, _LOCalc_RangeReplaceAll,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeReplace(ByRef $oRange, ByRef $oSrchDescript, $sSearchString, $sReplaceString)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iReplacements
	Local $oResult

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sReplaceString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSrchDescript.SearchString = $sSearchString
	$oSrchDescript.ReplaceString = $sReplaceString

	$oResult = $oRange.findFirst($oSrchDescript)
	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_SUCCESS, 1, 0) ; No Results

	$iReplacements = $oResult.replaceAll($oSrchDescript)

	Return ($iReplacements > 0) ? (SetError($__LO_STATUS_SUCCESS, 1, $oResult)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_RangeReplace

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeReplaceAll
; Description ...: Replace all instances of a search.
; Syntax ........: _LOCalc_RangeReplaceAll(ByRef $oRange, ByRef $oSrchDescript, $sSearchString, $sReplaceString)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a Regular Expression to Search for.
;                  $sReplaceString      - a string value. A String of text or a Regular Expression to replace any results with.
; Return values .: Success: 0 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;				   @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;				   @Error 1 @Extended 5 Return 0 = $sReplaceString not a String.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Results were found, but failed to perform the replacement.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Search and Replace was successful, @Extended set to number of replacements made, returning array Cell/CellRange Objects of all Cells modified.
;				   @Error 0 @Extended 0 Return 0 = Success. Search was successful, no results found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only the Sheet that contains the Range is searched, to search all Sheets you will have to cycle through and perform a search for each.
;				   Number of Replacements DOESN'T mean that is the size of the Array. If replacements where in several cells connected, the return will be a Cell Range for that area instead of individual cells.
; Related .......: _LOCalc_SearchDescriptorCreate, _LOCalc_RangeFindAll, _LOCalc_RangeFindNext, _LOCalc_RangeReplace
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeReplaceAll(ByRef $oRange, ByRef $oSrchDescript, $sSearchString, $sReplaceString)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResults
	Local $aoResults[0]
	Local $iReplacements

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sReplaceString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSrchDescript.SearchString = $sSearchString
	$oSrchDescript.ReplaceString = $sReplaceString

	$oResults = $oRange.findAll($oSrchDescript)
	If Not IsObj($oResults) Then Return SetError($__LO_STATUS_SUCCESS, 0, 0)

	If ($oResults.getCount() > 0) Then
		ReDim $aoResults[$oResults.getCount]
		For $i = 0 To $oResults.getCount() - 1
			$aoResults[$i] = $oResults.getByIndex($i)
			Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	$iReplacements = $oRange.replaceAll($oSrchDescript)

	Return ($iReplacements > 0) ? (SetError($__LO_STATUS_SUCCESS, $iReplacements, $aoResults)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_RangeReplaceAll

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowDelete
; Description ...: Delete Rows from a Sheet.
; Syntax ........: _LOCalc_RangeRowDelete(ByRef $oRange, $iRow[, $iCount = 1])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iRow                - an integer value. The Row to begin deleting at. The Row called will be deleted. See remarks.
;                  $iCount              - [optional] an integer value. Default is 1. The number of rows to delete, including the row called in $iRow.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iRow not an Integer or less than 0, or greater than number of Rows contained in the Range.
;				   @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested rows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Rows in L.O. Calc are 0 based, to Delete Row 1 in the LibreOffice UI, you would call $iRow with 0.
;				   Deleting Rows does not decrease the Row count, it simply erases the row's contents in a specific area and shifts all after content higher.
; Related .......: _LOCalc_RangeRowInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowDelete(ByRef $oRange, $iRow, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oRange.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iRow, 0, $oRows.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iCount, 1, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oRows.removeByIndex($iRow, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeRowDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowGetObjByPosition
; Description ...: Retrieve a Row's Object for further Row related functions.
; Syntax ........: _LOCalc_RangeRowGetObjByPosition(ByRef $oRange, $iRow)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iRow                - an integer value. The Row number to retrieve the Row Object for. See remarks.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iRow not an Integer or less than 0, or greater than number of Rows contained in the Range.
;				   @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Row Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Success, returning Row's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Rows in L.O. Calc are 0 based, to retrieve Row 1 in the LibreOffice UI, you would call $iRow with 0.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowGetObjByPosition(ByRef $oRange, $iRow)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows, $oRow

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oRange.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iRow, 0, $oRows.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRow = $oRows.getByIndex($iRow)
	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oRow)
EndFunc   ;==>_LOCalc_RangeRowGetObjByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowHeight
; Description ...: Set or Retrieve the Row's Height settings.
; Syntax ........: _LOCalc_RangeRowHeight(ByRef $oRow[, $bOptimal = Null[, $iHeight = Null]])
; Parameters ....: $oRow                - an object. A Row object returned by a previous _LOCalc_RangeRowGetObjByPosition function.
;                  $bOptimal            - [optional] a boolean value. Default is Null. If True, the Optimal height is automatically chosen.
;                  $iHeight             - [optional] an integer value (0-34464). Default is Null. The Height of the row, set in Micrometers.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bOptimal not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $iHeight not an Integer, less than 0 or greater than 34464.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $bOptimal
;				   |								2 = Error setting $iHeight
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: I am presently unable to find a setting for Optimal Height "Add" Value.
; Related .......: _LOCalc_ConvertFromMicrometer, _LOCalc_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowHeight(ByRef $oRow, $bOptimal = Null, $iHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avHeight[2]
	Local $iError = 0

	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOCalc_VarsAreNull($bOptimal, $iHeight) Then
		__LOCalc_ArrayFill($avHeight, $oRow.OptimalHeight(), $oRow.Height())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avHeight)
	EndIf

	If ($bOptimal <> Null) Then
		If Not IsBool($bOptimal) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oRow.OptimalHeight = $bOptimal
		$iError = ($oRow.OptimalHeight() = $bOptimal) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iHeight <> Null) Then
		If Not __LOCalc_IntIsBetween($iHeight, 0, 34464) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oRow.Height = $iHeight
		$iError = (__LOCalc_IntIsBetween($oRow.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeRowHeight

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowInsert
; Description ...: Insert blank rows from a specific row in a Range.
; Syntax ........: _LOCalc_RangeRowInsert(ByRef $oRange, $iRow[, $iCount = 1])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iRow                - an integer value. The Row to begin inserting blank rows at. See remarks. All contents from this row down will be shifted down.
;                  $iCount              - [optional] an integer value. Default is 1. The number of blank rows to insert.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iRow not an Integer or less than 0, or greater than number of Rows contained in the Range.
;				   @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully inserted blank rows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Rows in L.O. Calc are 0 based, to add Rows in Row 1 in the LibreOffice UI, you would call $iRow with 0.
;				   Inserting Rows does not increase the Row count, it simply adds blanks in a specific area and shifts all after content lower.
; Related .......: _LOCalc_RangeRowDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowInsert(ByRef $oRange, $iRow, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oRange.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iRow, 0, $oRows.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iCount, 1, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oRows.insertByIndex($iRow, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_RangeRowInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowPageBreak
; Description ...: Set or retrieve current Page Break settings for a Row.
; Syntax ........: _LOCalc_RangeRowPageBreak(ByRef $oRow[, $bManualPageBreak = Null[, $bStartOfPageBreak = Null]])
; Parameters ....: $oRow                - [in/out] an object. A Row object returned by a previous _LOCalc_RangeRowGetObjByPosition function.
;                  $bManualPageBreak    - [optional] a boolean value. Default is Null. If True, this row is the beginning of a manual Page Break.
;                  $bStartOfPageBreak   - [optional] a boolean value. Default is Null. If True, this row is the beginning of a start of Page Break. See Remarks.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRow not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bManualPageBreak not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bStartOfPageBreak not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $bManualPageBreak
;				   |								2 = Error setting $bStartOfPageBreak
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Setting $bStartOfPageBreak to True will insert a Manual Page Break, the same as setting $bManualPageBreak to True would.
;				   $bStartOfPageBreak setting is available more for indicating where Calc is inserting Page Breaks rather than for applying a setting. You can retrieve the settings for each row, and check if this value is set to True or not. If the Page break is an automatically inserted one, the value for $bManualPageBreak would be false.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowPageBreak(ByRef $oRow, $bManualPageBreak = Null, $bStartOfPageBreak = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abBreak[2]

	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOCalc_VarsAreNull($bManualPageBreak, $bStartOfPageBreak) Then
		__LOCalc_ArrayFill($abBreak, $oRow.IsManualPageBreak(), $oRow.IsStartOfNewPage())

		Return SetError($__LO_STATUS_SUCCESS, 1, $abBreak)
	EndIf

	If ($bManualPageBreak <> Null) Then
		If Not IsBool($bManualPageBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oRow.IsManualPageBreak = $bManualPageBreak
		$iError = ($oRow.IsManualPageBreak() = $bManualPageBreak) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bStartOfPageBreak <> Null) Then
		If Not IsBool($bStartOfPageBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oRow.IsStartOfNewPage = $bStartOfPageBreak
		$iError = ($oRow.IsStartOfNewPage() = $bStartOfPageBreak) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeRowPageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowsGetCount
; Description ...: Retrieve the total count of Rows contained in a Range.
; Syntax ........: _LOCalc_RangeRowsGetCount(ByRef $oRange)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Returning number of Rows contained in the Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: There is a fixed number of Rows per sheet, but different L.O. versions contain different amounts of Rows. This can also help determine how many rows are in a Cell Range.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowsGetCount(ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oRange.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oRows.Count())
EndFunc   ;==>_LOCalc_RangeRowsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_RangeRowVisible
; Description ...: Set or Retrieve the Row's visibility setting.
; Syntax ........: _LOCalc_RangeRowVisible(ByRef $oRow[, $bVisible = Null])
; Parameters ....: $oRow                - an object. A Row object returned by a previous _LOCalc_RangeRowGetObjByPosition function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the Row is Visible.
; Return values .: Success: 1 or Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRow not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $bVisible
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Boolean = Success. All optional parameters were set to Null, returning Row's current visibility setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_RangeRowVisible(ByRef $oRow, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bVisible = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oRow.IsVisible())


	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$oRow.IsVisible = $bVisible
	$iError = ($oRow.IsVisible() = $bVisible) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_RangeRowVisible
