#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Internal.au3"

; Other includes for Calc
#include "LibreOfficeCalc_Font.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for performing various Cursor movements.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_SheetCursorMove
; _LOCalc_TextCursorFont
; _LOCalc_TextCursorInsertString
; _LOCalc_TextCursorIsCollapsed
; _LOCalc_TextCursorMove
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetCursorMove
; Description ...: Move a Sheet Cursor object in a document. Also for creating/Expanding selections.
; Syntax ........: _LOCalc_SheetCursorMove(ByRef $oCursor, $iMove[, $iCount = 1[, $bSelect = False]])
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions.
;                  $iMove               - an integer value. The movement command. See remarks and Constants, $LOC_TEXTCUR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iColumns            - [optional] an integer value. Default is 0. The Number of Columns either to contain in the Range, or to move, depending on the called move command.
;                  $iRows               - [optional] an integer value. Default is 0. The Number of Rows either to contain in the Range, or to move, depending on the called move command.
;                  $iCount              - [optional] an integer value. Default is 1. Number of movements to make. See remarks.
;                  $bSelect             - [optional] a boolean value. Default is False. Whether to select data during this cursor movement. See remarks.
; Return values .: Success: 1.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iMove less than 0 or greater than highest move Constant. See Move Constants, $LOC_TEXTCUR_* as defined in LibreOfficeCalc_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $iColumns not an integer.
;				   @Error 1 @Extended 5 Return 0 = $iRows not an integer.
;				   @Error 1 @Extended 6 Return 0 = $iCount not an integer or is a negative.
;				   @Error 1 @Extended 7 Return 0 = $bSelect not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error determining cursor type.
;				   @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;				   @Error 3 @Extended 3 Return 0 = $oCursor Object unknown cursor type.
;				   --Success--
;				   @Error 0 @Extended ? Return 1 = Success, Cursor object movement was processed successfully. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only some movements accept Column and Row Values, creating/ extending a selection of cells, etc. They will be specified below.
;					#Cursor Movement Constants which accept Column and Row values:
;						$LOC_SHEETCUR_COLLAPSE_TO_SIZE, Call $iColumns with the number of columns to resize the range to contain, counting from the Top-Left hand cell of the current range. And call $iRows with the number of Rows to resize the range to contain, counting from the Top-Left hand cell of the current range..
;						$LOC_SHEETCUR_GOTO_OFFSET Call $iColumns with the number of columns to move, either left (negative number) or right (positive number), and call $iRows with the number of Rows to move up (negative number) or down (positive number).
;					#Cursor Movements which accept Selecting Only:
;						$LOC_SHEETCUR_GOTO_USED_AREA_START, Call $bSelect with a Boolean whether to select cells while performing this move (True) or not.
;						$LOC_SHEETCUR_GOTO_USED_AREA_END Call $bSelect with a Boolean whether to select cells while performing this move (True) or not.
;					#Cursor Movements which accept nothing and are done once per call:
;						$LOC_SHEETCUR_COLLAPSE_TO_CURRENT_ARRAY,
;						$LOC_SHEETCUR_COLLAPSE_TO_CURRENT_REGION,
;						$LOC_SHEETCUR_COLLAPSE_TO_MERGED_AREA,
;						$LOC_SHEETCUR_EXPAND_TO_ENTIRE_COLUMN,
;						$LOC_SHEETCUR_EXPAND_TO_ENTIRE_ROW,
;						$LOC_SHEETCUR_GOTO_START,
;						$LOC_SHEETCUR_GOTO_END
;					#Cursor Movements which accept only number of moves ($iCount):
;						$LOC_SHEETCUR_GOTO_NEXT, Call $iCount with the number of moves to perform.
;						$LOC_SHEETCUR_GOTO_PREV Call $iCount with the number of moves to perform.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOCalc_SheetCursorMove(ByRef $oCursor, $iMove, $iColumns = 0, $iRows = 0, $iCount = 1, $bSelect = False)
	Local $iCursorType
	Local $bMoved = False

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOCalc_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Switch $iCursorType
		Case $LOC_CURTYPE_SHEET_CURSOR
			$bMoved = __LOCalc_SheetCursorMove($oCursor, $iMove, $iColumns, $iRows, $iCount, $bSelect)
			Return SetError(@error, @extended, $bMoved)
		Case Else
			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; unknown or wrong cursor type.
	EndSwitch
EndFunc   ;==>_LOCalc_SheetCursorMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorFont
; Description ...: Set and Retrieve the Font Settings for a Text Cursor.
; Syntax ........: _LOCalc_TextCursorFont(ByRef $oDoc, ByRef $oTextCursor[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $sFontName           - [optional] a string value. Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] a general number value. Default is Null. The new Font size.
;                  $iPosture            - [optional] an integer value (0-5). Default is Null. The Font Italic setting. See Constants, $LOC_POSTURE_* as defined in LibreOfficeCalc_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] an integer value(0,50-200). Default is Null. The Font Bold settings see Constants, $LOC_WEIGHT_* as defined in LibreOfficeCalc_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oTextCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oTextCursor does not support Character properties.
;				   @Error 1 @Extended 4 Return 0 = Font called in $sFontName not available.
;				   @Error 1 @Extended 5 Return 0 = Variable passed to internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $sFontName not a String.
;				   @Error 1 @Extended 7 Return 0 = $nFontSize not a number.
;				   @Error 1 @Extended 8 Return 0 = $iPosture not an Integer, less than 0, or greater than 5. See Constants, $LOC_POSTURE_* as defined in LibreOfficeCalc_Constants.au3.
;				   @Error 1 @Extended 9 Return 0 = $iWeight not an Integer, less than 50 but not equal to 0, or greater than 200. See Constants, $LOC_WEIGHT_* as defined in LibreOfficeCalc_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $sFontName
;				   |								2 = Error setting $nFontSize
;				   |								4 = Error setting $iPosture
;				   |								8 = Error setting $iWeight
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted, such as oblique, ultra Bold etc.
;				   Libre Calc accepts only the predefined weight values, any other values are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......: _LOCalc_FontsList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorFont(ByRef $oDoc, ByRef $oTextCursor, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oTextCursor.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($sFontName <> Null) And Not _LOCalc_FontExists($oDoc, $sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOCalc_CellFont($oTextCursor, $sFontName, $nFontSize, $iPosture, $iWeight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_TextCursorFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorInsertString
; Description ...: Insert a string using a Text cursor.
; Syntax ........: _LOCalc_TextCursorInsertString(ByRef $oCursor, $sString[, $bOverwrite = False])
; Parameters ....: $oCursor             - [in/out] an object. A Text Cursor Object returned from any Cursor Object creation function.
;                  $sString             - a string value. A String to insert.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, and the cursor object has text selected, the selection is overwritten, else if False, the string is inserted to the left of the selection.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sString not a string..
;				   @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $oCursor is not a Text Cursor.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Cursor type.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. String was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_PageStyleHeaderCreateTextCursor, _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_CellCreateTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorInsertString(ByRef $oCursor, $sString, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	$iCursorType = __LOCalc_Internal_CursorGetType($oCursor)
	If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iCursorType <> $LOC_CURTYPE_TEXT_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oCursor.Text.insertString($oCursor, $sString, $bOverwrite)
	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_TextCursorInsertString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorIsCollapsed
; Description ...: Retrieve the current status of a Text cursor, whether it has any data selected or not.
; Syntax ........: _LOCalc_TextCursorIsCollapsed(ByRef $oCursor)
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor unknown cursor type.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean. = Success. Successfully queried whether cursor's selection is collapsed. Returning Boolean result, True = Cursor has no data selected, False = cursor has data selected.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_TextCursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorIsCollapsed(ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $bReturn

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOCalc_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Switch $iCursorType
		Case $LOC_CURTYPE_TEXT_CURSOR
			$bReturn = $oCursor.isCollapsed()
			Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)

		Case Else
			Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; unknown or wrong cursor data type.
	EndSwitch
EndFunc   ;==>_LOCalc_TextCursorIsCollapsed

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorMove
; Description ...: Move a Text Cursor object in a document. Also for creating/Expanding selections.
; Syntax ........: _LOCalc_TextCursorMove(ByRef $oCursor, $iMove[, $iCount = 1[, $bSelect = False]])
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions.
;                  $iMove               - an integer value. The movement command. See remarks and Constants, $LOC_TEXTCUR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iCount              - [optional] an integer value. Default is 1. Number of movements to make. See remarks.
;                  $bSelect             - [optional] a boolean value. Default is False. Whether to select data during this cursor movement. See remarks.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iMove less than 0 or greater than highest move Constant. See Move Constants, $LOC_TEXTCUR_* as defined in LibreOfficeCalc_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $iCount not an integer or is a negative.
;				   @Error 1 @Extended 5 Return 0 = $bSelect not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error determining cursor type.
;				   @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;				   @Error 3 @Extended 3 Return 0 = $oCursor Object unknown cursor type.
;				   --Success--
;				   @Error 0 @Extended ? Return Boolean = Success, Cursor object movement was processed successfully. Returns True if the full count of movements were successful, else false if none or only partially successful. @Extended set to number of successful movements. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only some movements accept movement amounts and selecting (such as $LOC_TEXTCUR_GO_RIGHT 2, True) etc. Also only some accept creating/ extending a selection of text/ data. They will be specified below.
;					 To Clear /Unselect a current selection, you can input a move such as $LOC_TEXTCUR_GO_RIGHT, 0, False.
;					#Cursor Movement Constants which accept Number of Moves and Selecting:
;						$LOC_TEXTCUR_GO_LEFT,
;						$LOC_TEXTCUR_GO_RIGHT,
;					#Cursor Movements which accept Selecting Only:
;						$LOC_TEXTCUR_GOTO_START,
;						$LOC_TEXTCUR_GOTO_END,
;					#Cursor Movements which accept nothing and are done once per call:
;						$LOC_TEXTCUR_COLLAPSE_TO_START,
;						$LOC_TEXTCUR_COLLAPSE_TO_END
; Related .......: _LOCalc_TextCursorIsCollapsed
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorMove(ByRef $oCursor, $iMove, $iCount = 1, $bSelect = False)
	Local $iCursorType
	Local $bMoved = False

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOCalc_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Switch $iCursorType
		Case $LOC_CURTYPE_TEXT_CURSOR
			$bMoved = __LOCalc_TextCursorMove($oCursor, $iMove, $iCount, $bSelect)
			Return SetError(@error, @extended, $bMoved)
		Case Else
			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; unknown or wrong cursor type.
	EndSwitch
EndFunc   ;==>_LOCalc_TextCursorMove
