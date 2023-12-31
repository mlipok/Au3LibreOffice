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
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Removing, etc. L.O. Calc document Sheets.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_SheetActivate
; _LOCalc_SheetAdd
; _LOCalc_SheetColumnDelete
; _LOCalc_SheetColumnGetName
; _LOCalc_SheetColumnGetObjByName
; _LOCalc_SheetColumnGetObjByPosition
; _LOCalc_SheetColumnInsert
; _LOCalc_SheetColumnPageBreak
; _LOCalc_SheetColumnsGetCount
; _LOCalc_SheetColumnVisible
; _LOCalc_SheetColumnWidth
; _LOCalc_SheetCopy
; _LOCalc_SheetGetActive
; _LOCalc_SheetGetCellByName
; _LOCalc_SheetGetCellByPosition
; _LOCalc_SheetGetObjByName
; _LOCalc_SheetIsActive
; _LOCalc_SheetMove
; _LOCalc_SheetName
; _LOCalc_SheetRemove
; _LOCalc_SheetRowDelete
; _LOCalc_SheetRowGetObjByPosition
; _LOCalc_SheetRowHeight
; _LOCalc_SheetRowInsert
; _LOCalc_SheetRowPageBreak
; _LOCalc_SheetRowsGetCount
; _LOCalc_SheetRowVisible
; _LOCalc_SheetsGetCount
; _LOCalc_SheetsGetNames
; _LOCalc_SheetVisible
; ===============================================================================================================================





; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetActivate
; Description ...: Activate a Sheet in a Calc Document.
; Syntax ........: _LOCalc_SheetActivate(ByRef $oDoc, ByRef $oSheet)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Sheet was successfully activated.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:_LOCalc_SheetIsActive
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetActivate(ByRef $oDoc, ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.setActiveSheet($oSheet)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetActivate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetAdd
; Description ...: Insert a new Sheet into a Calc Document.
; Syntax ........: _LOCalc_SheetAdd(ByRef $oDoc[, $sName = Null[, $iPosition = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sName               - [optional] a string value. Default is Null. The Name of the new Sheet. See remarks.
;                  $iPosition           - [optional] an integer value. Default is Null. The position to insert the new sheet. If left as Null, new sheet is inserted at the end. See remarks.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sName not a String.
;				   @Error 1 @Extended 3 Return 0 = Document already contains a Sheet named the same as called in $sName.
;				   @Error 1 @Extended 4 Return 0 = $iPosition not an Integer, less than 0 or greater than number of sheets present in the document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve new Sheet's Object. New Sheet may not have been inserted successfully.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. New sheet was successfully inserted, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $sName is left as Null, the sheet will be automatically named "Sheet?" where "?" is a digit.
;				   If $iPosition is left as Null, the sheet will be inserted at the end of the list of Sheets.
;				   Calling $iPosition with the number of Sheets in the Document will place the added sheet at the end of the sheet list.
; Related .......: _LOCalc_SheetRemove, _LOCalc_DocHasSheetName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetAdd(ByRef $oDoc, $sName = Null, $iPosition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets, $oSheet
	Local $iCount = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sName = Null) Then
		$sName = "Sheet" & ($oSheets.Count() + 1)
		If $oSheets.hasByName($sName) Then
			$iCount = $oSheets.Count()
			While $oSheets.hasByName($sName)
				$iCount += 1
				$sName = "Sheet" & $iCount

			WEnd

		EndIf

	EndIf

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oSheets.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($iPosition = Null) Then $iPosition = $oSheets.Count()

	If Not __LOCalc_IntIsBetween($iPosition, 0, $oSheets.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oSheets.insertNewByName($sName, $iPosition)

	$oSheet = $oSheets.getByName($sName)

	Return (IsObj($oSheet)) ? (SetError($__LO_STATUS_SUCCESS, 0, $oSheet)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_SheetAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetColumnDelete
; Description ...: Delete Columns from a Sheet.
; Syntax ........: _LOCalc_SheetColumnDelete(ByRef $oSheet, $iColumn[, $iCount = 1])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iColumn             - an integer value. The column to begin deleting at. The Column called will be deleted. See remarks.
;                  $iCount              - [optional] an integer value. Default is 1. The number of columns to delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iColumns not an Integer or less than 0, or greater than number of Columns contained in the Sheet.
;				   @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested columns.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Columns in L.O. Calc are 0 based, to Delete Column "A" in the LibreOffice UI, you would call $iColumn with 0.
;				   Deleting Columns does not decrease the Column count, it simply erases the Column's contents in a specific area and shifts all after content left.
; Related .......: _LOCalc_SheetColumnInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetColumnDelete(ByRef $oSheet, $iColumn, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oSheet.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iColumn, 0, $oColumns.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iCount, 1, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumns.removeByIndex($iColumn, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetColumnDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetColumnGetName
; Description ...: Retrieve the Column's name.
; Syntax ........: _LOCalc_SheetColumnGetName(ByRef $oColumn)
; Parameters ....: $oColumn             - [in/out] an object. A Column object returned by a previous _LOCalc_SheetColumnGetObjByPosition, or _LOCalc_SheetColumnGetObjByName function.
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
Func _LOCalc_SheetColumnGetName(ByRef $oColumn)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sName = $oColumn.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sName)
EndFunc   ;==>_LOCalc_SheetColumnGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetColumnGetObjByName
; Description ...: Retrieve a Column's Object by name.
; Syntax ........: _LOCalc_SheetColumnGetObjByName(ByRef $oSheet, $sName)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sName               - a string value. The Column name to retrieve the Object for, such as "A".
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sName not a String.
;				   @Error 1 @Extended 3 Return 0 = Sheet does not contain a column with name called in $sName.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Column Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Success, returning Column's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetColumnGetObjByPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetColumnGetObjByName(ByRef $oSheet, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns, $oColumn

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumns = $oSheet.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not $oColumns.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumn = $oColumns.getByName($sName)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOCalc_SheetColumnGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetColumnGetObjByPosition
; Description ...: Retrieve the Column's Object by its position.
; Syntax ........: _LOCalc_SheetColumnGetObjByPosition(ByRef $oSheet, $iColumn)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iColumn             - an integer value.The Column number to retrieve the Object for. See remarks.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, or less than 0, or greater than number of columns contained in the Sheet.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Column Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Success, returning Column's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Columns in L.O. Calc are 0 based, to retrieve Column "A" in the LibreOffice UI, you would call $iColumn with 0.
; Related .......: _LOCalc_SheetColumnGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetColumnGetObjByPosition(ByRef $oSheet, $iColumn)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns, $oColumn

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oSheet.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iColumn, 0, $oColumns.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumn = $oColumns.getByIndex($iColumn)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOCalc_SheetColumnGetObjByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetColumnInsert
; Description ...: Insert blank columns into a sheet as a specific column.
; Syntax ........: _LOCalc_SheetColumnInsert(ByRef $oSheet, $iColumn[, $iCount = 1])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iColumn             - an integer value. The Column to insert the new column(s) at. See remarks. New columns will be inserted starting at this column and all content will be shifted right.
;                  $iCount              - [optional] an integer value. Default is 1. The number of blank columns to insert.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iColumn not an Integer or less than 0, or greater than number of Columns contained in the Sheet.
;				   @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully inserted blank columns.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Columns in L.O. Calc are 0 based, to add columns in Column "A" in the LibreOffice UI, you would call $iColumn with 0.
;				   Inserting Columnss does not increase the Column count, it simply adds blanks in a specific area and shifts all after content further right.
; Related .......: _LOCalc_SheetColumnDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetColumnInsert(ByRef $oSheet, $iColumn, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oSheet.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iColumn, 0, $oColumns.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iCount, 1, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumns.insertByIndex($iColumn, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetColumnInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetColumnPageBreak
; Description ...: Set or retrieve current Page Break settings for a Column.
; Syntax ........: _LOCalc_SheetColumnPageBreak(ByRef $oColumn[, $bManualPageBreak = Null[, $bStartOfPageBreak = Null]])
; Parameters ....: $oColumn             - [in/out] an object. A Column object returned by a previous _LOCalc_SheetColumnGetObjByPosition, or _LOCalc_SheetColumnGetObjByName function.
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
Func _LOCalc_SheetColumnPageBreak(ByRef $oColumn, $bManualPageBreak = Null, $bStartOfPageBreak = Null)
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
EndFunc   ;==>_LOCalc_SheetColumnPageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetColumnsGetCount
; Description ...: Retrieve the total count of Columns contained in a Sheet.
; Syntax ........: _LOCalc_SheetColumnsGetCount(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Returning number of Columns contained in the Sheet.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: There is a fixed number of Columns per sheet, but different L.O. versions contain different amounts of Columns.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOCalc_SheetColumnsGetCount(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oSheet.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumns.Count())
EndFunc   ;==>_LOCalc_SheetColumnsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetColumnVisible
; Description ...: Set or Retrieve the Column's visibility setting.
; Syntax ........: _LOCalc_SheetColumnVisible($oColumn[, $bVisible = Null])
; Parameters ....: $oColumn             - an object. A Column object returned by a previous _LOCalc_SheetColumnGetObjByPosition, or _LOCalc_SheetColumnGetObjByName function.
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
Func _LOCalc_SheetColumnVisible($oColumn, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bVisible = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oColumn.IsVisible())


	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$oColumn.IsVisible = $bVisible
	$iError = ($oColumn.IsVisible() = $bVisible) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_SheetColumnVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetColumnWidth
; Description ...: Set or Retrieve the Column's Width settings.
; Syntax ........: _LOCalc_SheetColumnWidth($oColumn[, $bOptimal = Null[, $iWidth = Null]])
; Parameters ....: $oColumn             - an object. A Column object returned by a previous _LOCalc_SheetColumnGetObjByPosition, or _LOCalc_SheetColumnGetObjByName function.
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
Func _LOCalc_SheetColumnWidth($oColumn, $bOptimal = Null, $iWidth = Null)
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
EndFunc   ;==>_LOCalc_SheetColumnWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetCopy
; Description ...: Create a copy of a particular Sheet.
; Syntax ........: _LOCalc_SheetCopy(ByRef $oDoc, ByRef $oSheet, $sNewName, $iPosition)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sNewName            - a string value. The name to assign to the newly copied Sheet.
;                  $iPosition           - an integer value. The position to place the copied sheet at. 0 = the beginning.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sNewName not a String.
;				   @Error 1 @Extended 4 Return 0 = Document already contains a Sheet with the same name as called in $sNewName.
;				   @Error 1 @Extended 5 Return 0 = $iPosition not an Integer, less than 0, or greater than number of Sheets contained in the document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve original Sheet's name.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve Object for new Sheet.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully copied the Sheet. Returning the new Sheet's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $sNewName is left as Null, the original Sheet's name is used, with "_" and a digit appended.
;				   If $iPosition is left as Null, the copied sheet will be placed at the end of the list.
;				   Calling $iPosition with the number of Sheets in the Document will place the copied sheet at the end of the sheet list.
; Related .......: _LOCalc_DocHasSheetName, _LOCalc_SheetMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetCopy(ByRef $oDoc, ByRef $oSheet, $sNewName = Null, $iPosition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets, $oNewSheet
	Local $sName
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$sName = $oSheet.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If ($sNewName = Null) Then
		$sNewName = $sName & "_" & 2
		If $oSheets.hasByName($sNewName) Then
			$iCount = 2
			While $oSheets.hasByName($sNewName)
				$iCount += 1
				$sNewName = $sName & "_" & $iCount

			WEnd

		EndIf

	EndIf

	If Not IsString($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If $oSheets.hasByName($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If ($iPosition = Null) Then $iPosition = $oSheets.Count()

	If Not __LOCalc_IntIsBetween($iPosition, 0, $oSheets.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)



	$oSheets.copyByName($sName, $sNewName, $iPosition)

	$oNewSheet = $oSheets.getByName($sNewName)
	If Not IsObj($oNewSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNewSheet)
EndFunc   ;==>_LOCalc_SheetCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetGetActive
; Description ...: Retrieve a Sheet object for the currently active Sheet.
; Syntax ........: _LOCalc_SheetGetActive(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Active Sheet's Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully retrieved the Active Sheet, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetGetActive(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheet = $oDoc.CurrentController.getActiveSheet()
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheet)
EndFunc   ;==>_LOCalc_SheetGetActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetGetCellByName
; Description ...: Retrieve a Cell or Cell Range Object by Cell name.
; Syntax ........: _LOCalc_SheetGetCellByName(ByRef $oSheet, $sFromCellName[, $sToCellName = Null])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sFromCellName       - a string value. The cell to retrieve the Object for, or to begin the Cell Range. See remarks.
;                  $sToCellName         - [optional] a string value. Default is Null. The cell to end the Cell Range at.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFromCellName not a String.
;				   @Error 1 @Extended 3 Return 0 = $sToCellName not set to Null, and not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve requested Cell or Cell Range Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully retrieved and returning requested Cell or Cell Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFromCellName can be a Cell Name or a defined Cell Range name.
; Related .......: _LOCalc_SheetGetCellByPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetGetCellByName(ByRef $oSheet, $sFromCellName, $sToCellName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCellRange
	Local $sCellRange = $sFromCellName

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFromCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($sToCellName <> Null) And Not IsString($sToCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sToCellName <> Null) Then $sCellRange &= ":" & $sToCellName

	$oCellRange = $oSheet.getCellRangeByName($sCellRange)
	If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCellRange)
EndFunc   ;==>_LOCalc_SheetGetCellByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetGetCellByPosition
; Description ...: Retrieve a Cell or Cell Range Object by position.
; Syntax ........: _LOCalc_SheetGetCellByPosition(ByRef $oSheet, $iColumn, $iRow[, $iToColumn = Null[, $iToRow = Null]])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iColumn             - an integer value. The Column of the desired cell, or of the beginning of the Cell range. 0 Based. See remarks.
;                  $iRow                - an integer value. The Row of the desired cell, or of the beginning of the Cell range. 0 Based. See remarks.
;                  $iToColumn           - [optional] an integer value. Default is Null. The Column of the end of the Cell range. 0 Based. Must be greater or equal to $iColumn.
;                  $iToRow              - [optional] an integer value. Default is Null. The Row of the end of the Cell range. 0 Based. Must be greater or equal to $iRow.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, or less than 0.
;				   @Error 1 @Extended 3 Return 0 = $iRow not an Integer, or less than 0.
;				   @Error 1 @Extended 4 Return 0 = $iToColumn not an Integer, or less than 0.
;				   @Error 1 @Extended 5 Return 0 = $iToRow not an Integer, or less than 0.
;				   @Error 1 @Extended 6 Return 0 = $iToColumn less than $iColumn.
;				   @Error 1 @Extended 7 Return 0 = $iToRow less than $iRow.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve an individual Cell's Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve a Cell Range's Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully retrieved and returned an Individual Cell's Object.
;				   @Error 0 @Extended 1 Return Object = Success. Successfully retrieved and returned a Cell Range's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: According to the wiki (https://wiki.documentfoundation.org/Faq/Calc/022), the maximum Columns contained in a sheet is 1024 until version 7.3, or 16384 from 7.3. and up..
;				   According to Andrew Pitonyak, (OOME. 4.1 Page 492), the maximum number of rows contained in a sheet is 65,536 as of OOo Calc 3.0, but according to the wiki (https://wiki.documentfoundation.org/Faq/Calc/022), the maximum or Rows for Libre Office Calc is 1,048,576.
; Related .......: _LOCalc_SheetGetCellByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetGetCellByPosition(ByRef $oSheet, $iColumn, $iRow, $iToColumn = Null, $iToRow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell, $oCellRange

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iColumn, 0, $iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iRow, 0, $iRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($iToColumn <> Null) Or ($iToRow <> Null) Then
		If Not __LOCalc_IntIsBetween($iToColumn, 0, $iToColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If Not __LOCalc_IntIsBetween($iToRow, 0, $iToRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If ($iToColumn < $iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If ($iToRow < $iRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	EndIf

	If ($iToColumn = Null) And ($iToRow = Null) Then
		$oCell = $oSheet.getCellByPosition($iColumn, $iRow)
		If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		Return SetError($__LO_STATUS_SUCCESS, 0, $oCell)

	Else
		$oCellRange = $oSheet.getCellRangeByPosition($iColumn, $iRow, $iToColumn, $iToRow)
		If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
		Return SetError($__LO_STATUS_SUCCESS, 1, $oCellRange)

	EndIf
EndFunc   ;==>_LOCalc_SheetGetCellByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetGetObjByName
; Description ...: Retrieve a Sheet Object for a specific Sheet by name.
; Syntax ........: _LOCalc_SheetGetObjByName(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sName               - a string value. The sheet name to retrieve the Object for.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sName not a String.
;				   @Error 1 @Extended 3 Return 0 = Document does not contain a sheet with name called in $sName.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve requested Sheet's object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning requested Sheet's object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetGetObjByName(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet, $oSheets

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not $oSheets.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSheet = $oSheets.getByName($sName)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheet)
EndFunc   ;==>_LOCalc_SheetGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetIsActive
; Description ...: Check if a particular Sheet is the active Sheet.
; Syntax ........: _LOCalc_SheetIsActive(ByRef $oDoc, ByRef $oSheet)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. If the called Sheet is the currently active sheet, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetActivate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetIsActive(ByRef $oDoc, ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheet2

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheet2 = $oDoc.CurrentController.getActiveSheet()
	If Not IsObj($oSheet2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($oSheet.AbsoluteName() = $oSheet2.AbsoluteName()) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOCalc_SheetIsActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetMove
; Description ...: Move a Sheet's position in the list of Sheets in a Calc Document.
; Syntax ........: _LOCalc_SheetMove(ByRef $oDoc, ByRef $oSheet, $iPosition)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iPosition           - an integer value. The Position the move the Sheet to, 0 being the beginning.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iPosition not an Integer, less than 0 or greater than number of sheets contained in the document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Sheet's name.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Sheet was successfully moved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling $iPosition with the number of Sheets in the Document will place the moved sheet at the end of the sheet list.
; Related .......: _LOCalc_SheetCopy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetMove(ByRef $oDoc, ByRef $oSheet, $iPosition)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets
	Local $sName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)


	$sName = $oSheet.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not __LOCalc_IntIsBetween($iPosition, 0, $oSheets.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSheets.moveByName($sName, $iPosition)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetName
; Description ...: Set or Retrieve a Sheet's name.
; Syntax ........: _LOCalc_SheetName(ByRef $oDoc, ByRef $oSheet[, $sName = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $sName               - [optional] a string value. Default is Null. The new name for the Sheet.
; Return values .: Success: 1 or String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sName not a String.
;				   @Error 1 @Extended 4 Return 0 = Document already has a Sheet named the same as called in $sName.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $sName
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Sheet's new name was successfully set.
;				   @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning the Sheet's current name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOCalc_DocHasSheetName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetName(ByRef $oDoc, ByRef $oSheet, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($sName = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oSheet.Name())

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $oSheets.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oSheet.Name = $sName

	$iError = ($oSheet.Name() = $sName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_SheetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetRemove
; Description ...: Remove a Sheet from a Calc Document.
; Syntax ........: _LOCalc_SheetRemove(ByRef $oDoc, ByRef $oSheet)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSheet not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve the Sheet's name.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Sheets Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to delete the Sheet, but a Sheet by that name still exists.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully removed the requested sheet.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocSheetAdd, _LOCalc_SheetGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetRemove(ByRef $oDoc, ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets
	Local $sName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sName = $oSheet.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oSheets.removeByName($sName)

	If $oSheets.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetRemove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetRowDelete
; Description ...: Delete Rows from a Sheet.
; Syntax ........: _LOCalc_SheetRowDelete(ByRef $oSheet, $iRow[, $iCount = 1])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iRow                - an integer value. The Row to begin deleteing at. The Row called will be deleted. See remarks.
;                  $iCount              - [optional] an integer value. Default is 1. The number of rows to delete, including the row called in $iRow.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iRow not an Integer or less than 0, or greater than number of Rows contained in the Sheet.
;				   @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested rows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Rows in L.O. Calc are 0 based, to Delete Row 1 in the LibreOffice UI, you would call $iRow with 0.
;				   Deleting Rows does not decrease the Row count, it simply erases the row's contents in a specific area and shifts all after content higher.
; Related .......: _LOCalc_SheetRowInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetRowDelete(ByRef $oSheet, $iRow, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oSheet.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iRow, 0, $oRows.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iCount, 1, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oRows.removeByIndex($iRow, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetRowDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetRowGetObjByPosition
; Description ...: Retrieve a Row's Object for further Row related functions.
; Syntax ........: _LOCalc_SheetRowGetObjByPosition(ByRef $oSheet, $iRow)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iRow                - an integer value. The Row number to retrieve the Row Object for. See remarks.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iRow not an Integer or less than 0, or greater than number of Rows contained in the Sheet.
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
Func _LOCalc_SheetRowGetObjByPosition(ByRef $oSheet, $iRow)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows, $oRow

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oSheet.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iRow, 0, $oRows.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRow = $oRows.getByIndex($iRow)
	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oRow)
EndFunc   ;==>_LOCalc_SheetRowGetObjByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetRowHeight
; Description ...: Set or Retrieve the Row's Height settings.
; Syntax ........: _LOCalc_SheetRowHeight($oRow[, $bOptimal = Null[, $iHeight = Null]])
; Parameters ....: $oRow                - an object. A Row object returned by a previous _LOCalc_SheetRowGetObjByPosition function.
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
Func _LOCalc_SheetRowHeight($oRow, $bOptimal = Null, $iHeight = Null)
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
EndFunc   ;==>_LOCalc_SheetRowHeight

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetRowInsert
; Description ...: Insert blank rows from a specific row in a sheet.
; Syntax ........: _LOCalc_SheetRowInsert(ByRef $oSheet, $iRow[, $iCount = 1])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $iRow                - an integer value. The Row to begin inserting blank rows at. See remarks. All contents from this row down will be shifted down.
;                  $iCount              - [optional] an integer value. Default is 1. The number of blank rows to insert.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iRow not an Integer or less than 0, or greater than number of Rows contained in the Sheet.
;				   @Error 1 @Extended 3 Return 0 = $iCount not an Integer, or less than 1.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully inserted blank rows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Rows in L.O. Calc are 0 based, to add Rows in Row 1 in the LibreOffice UI, you would call $iRow with 0.
;				   Inserting Rows does not increase the Row count, it simply adds blanks in a specific area and shifts all after content lower.
; Related .......: _LOCalc_SheetRowDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetRowInsert(ByRef $oSheet, $iRow, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oSheet.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not __LOCalc_IntIsBetween($iRow, 0, $oRows.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iCount, 1, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oRows.insertByIndex($iRow, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SheetRowInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetRowPageBreak
; Description ...: Set or retrieve current Page Break settings for a Row.
; Syntax ........: _LOCalc_SheetRowPageBreak(ByRef $oRow[, $bManualPageBreak = Null[, $bStartOfPageBreak = Null]])
; Parameters ....: $oRow                - [in/out] an object. A Row object returned by a previous _LOCalc_SheetRowGetObjByPosition function.
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
Func _LOCalc_SheetRowPageBreak(ByRef $oRow, $bManualPageBreak = Null, $bStartOfPageBreak = Null)
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
EndFunc   ;==>_LOCalc_SheetRowPageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetRowsGetCount
; Description ...: Retrieve the total count of Rows contained in a Sheet.
; Syntax ........: _LOCalc_SheetRowsGetCount(ByRef $oSheet)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Rows Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Returning number of Rows contained in the Sheet.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: There is a fixed number of Rows per sheet, but different L.O. versions contain different amounts of Rows.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOCalc_SheetRowsGetCount(ByRef $oSheet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRows

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRows = $oSheet.getRows()
	If Not IsObj($oRows) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oRows.Count())
EndFunc   ;==>_LOCalc_SheetRowsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetRowVisible
; Description ...: Set or Retrieve the Row's visibility setting.
; Syntax ........: _LOCalc_SheetRowVisible($oRow[, $bVisible = Null])
; Parameters ....: $oRow                - an object. A Row object returned by a previous _LOCalc_SheetRowGetObjByPosition function.
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
Func _LOCalc_SheetRowVisible($oRow, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bVisible = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oRow.IsVisible())


	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$oRow.IsVisible = $bVisible
	$iError = ($oRow.IsVisible() = $bVisible) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_SheetRowVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetsGetCount
; Description ...: Retrieve a count of Sheets contained in a Calc Document.
; Syntax ........: _LOCalc_SheetsGetCount(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Returning count of Sheets contained in the Calc Document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetsGetCount(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSheets.Count())
EndFunc   ;==>_LOCalc_SheetsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetsGetNames
; Description ...: Retrieve an array of Sheet names for a Calc Document.
; Syntax ........: _LOCalc_SheetsGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning Array of Sheet names for this document. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_SheetGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetsGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets
	Local $asNames[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	ReDim $asNames[$oSheets.Count()]

	For $i = 0 To $oSheets.Count() - 1
		$asNames[$i] = $oSheets.getByIndex($i).Name()

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOCalc_SheetsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetVisible
; Description ...: Set or Retrieve a Sheet's current visibility setting.
; Syntax ........: _LOCalc_SheetVisible(ByRef $oSheet[, $bVisible = Null])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the Sheet is visible in the Libre Office UI.
; Return values .:  Success: 1 or Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $bVisiblee
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Sheet Visibility setting was successfully set.
;				   @Error 0 @Extended 1 Return Boolean = Success. $bVisible set to Null, returning current visibility setting. True indicates the Sheet is currently visible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetVisible(ByRef $oSheet, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bVisible = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oSheet.IsVisible())

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheet.IsVisible = $bVisible

	Return ($oSheet.IsVisible = $bVisible) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_SheetVisible
