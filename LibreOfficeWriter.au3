;~ #AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Internal.au3"
#include "LibreOfficeWriter_Helper.au3"

; include in order like they was splitted from LibreOfficeWriter.au3
#include "LibreOfficeWriter_Doc.au3"
#include "LibreOfficeWriter_Frame.au3"
#include "LibreOfficeWriter_Table.au3"
#include "LibreOfficeWriter_DirectFormating.au3"
#include "LibreOfficeWriter_Field.au3"

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter)
; AutoIt Version : v3.3.16.1
; UDF Version    : 0.0.0.3
; Description ...: Provides basic functionality through Autoit for interacting with Libre Office Writer.
; Author(s) .....: donnyh13
; Sources .......: jguinch -- Printmgr.au3, used (_PrintMgr_EnumPrinter);
;					mLipok -- OOoCalc.au3, used (__OOoCalc_ComErrorHandler_UserFunction,_InternalComErrorHandler,
;						-- WriterDemo.au3, used _CreateStruct;
;					Andrew Pitonyak & Laurent Godard (VersionGet);
;					Leagnus & GMK -- OOoCalc.au3, used (SetPropertyValue)
; Dll ...........:
; Note...........: Tips/templates taken from OOoCalc UDF written by user GMK; also from Word UDF by user water.
;					I found the book by Andrew Pitonyak very helpful also, titled, "OpenOffice.org Macros Explained;
;						OOME Third Edition".
;					Of course, this UDF is written using the English version of LibreOffice, and may only work for the English
;						version of LibreOffice installations. Many functions in this UDF may or may not work with OpenOffice
;						Writer, however some settings are definitely for LibreOffice only.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_CellBackColor
; _LOWriter_CellBorderColor
; _LOWriter_CellBorderPadding
; _LOWriter_CellBorderStyle
; _LOWriter_CellBorderWidth
; _LOWriter_CellCreateTextCursor
; _LOWriter_CellFormula
; _LOWriter_CellGetDataType
; _LOWriter_CellGetError
; _LOWriter_CellGetName
; _LOWriter_CellProtect
; _LOWriter_CellString
; _LOWriter_CellValue
; _LOWriter_CellVertOrient
; _LOWriter_EndnoteDelete
; _LOWriter_EndnoteGetAnchor
; _LOWriter_EndnoteGetTextCursor
; _LOWriter_EndnoteInsert
; _LOWriter_EndnoteModifyAnchor
; _LOWriter_EndnoteSettingsAutoNumber
; _LOWriter_EndnoteSettingsStyles
; _LOWriter_EndnotesGetList
; _LOWriter_FootnoteDelete
; _LOWriter_FootnoteGetAnchor
; _LOWriter_FootnoteGetTextCursor
; _LOWriter_FootnoteInsert
; _LOWriter_FootnoteModifyAnchor
; _LOWriter_FootnoteSettingsAutoNumber
; _LOWriter_FootnoteSettingsContinuation
; _LOWriter_FootnoteSettingsStyles
; _LOWriter_FootnotesGetList
; _LOWriter_SearchDescriptorCreate
; _LOWriter_SearchDescriptorModify
; _LOWriter_SearchDescriptorSimilarityModify
; _LOWriter_ShapesGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBackColor
; Description ...: Set and Retrieve the Background color of a Cell or Cell Range.
; Syntax ........: _LOWriter_CellBackColor(Byref $oCell[, $iBackColor = Null[, $bBackTransparent = Null]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from any Table Cell Object creation or retrieval functions.
;                  $iBackColor          - [optional] an integer value. Default is Null. Specify the Cell background color as a Long Integer. See Remarks.
;				   +						Set to $LOW_COLOR_OFF(-1) to disable Background color.
;                  $bBackTransparent    - [optional] a boolean value. Default is Null. If True, the background color is transparent.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell variable not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iBackColor not an Integer, set less than -1 or greater than 16777215.
;				   @Error 1 @Extended 3 Return 0 = $bBackTransparent not a Boolean and not set to Null.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iBackColor
;				   |								2 = Error setting $bBackTransparent
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it. $iBackColor is set using Long integer.
;					See _LOWriter_ConvertColorToLong, _LOWriter_ConvertColorFromLong. There are also preset colors, listed below.
; Color Constants: $LOW_COLOR_OFF(-1)
;					$LOW_COLOR_BLACK(0),
;					$LOW_COLOR_WHITE(16777215),
;					$LOW_COLOR_LGRAY(11711154),
;					$LOW_COLOR_GRAY(8421504),
;					$LOW_COLOR_DKGRAY(3355443),
;					$LOW_COLOR_YELLOW(16776960),
;					$LOW_COLOR_GOLD(16760576),
;					$LOW_COLOR_ORANGE(16744448),
;					$LOW_COLOR_BRICK(16728064),
;					$LOW_COLOR_RED(16711680),
;					$LOW_COLOR_MAGENTA(12517441),
;					$LOW_COLOR_PURPLE(8388736),
;					$LOW_COLOR_INDIGO(5582989),
;					$LOW_COLOR_BLUE(2777241),
;					$LOW_COLOR_TEAL(1410150),
;					$LOW_COLOR_GREEN(43315),
;					$LOW_COLOR_LIME(8508442),
;					$LOW_COLOR_BROWN(9127187).
; Related .......: _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellBackColor(ByRef $oCell, $iBackColor = Null, $bBackTransparent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avColor[2]

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iBackColor, $bBackTransparent) Then
		__LOWriter_ArrayFill($avColor, $oCell.BackColor(), $oCell.BackTransparent())
		Return SetError($__LOW_STATUS_SUCCESS, 0, $avColor)
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iBackColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oCell.BackColor = $iBackColor
		If ($iBackColor = $LOW_COLOR_OFF) Then $oCell.BackTransparent = True
		$iError = ($oCell.BackColor() = $iBackColor) ? $iError : BitOR($iError, 1) ; Error setting color.
	EndIf

	If ($bBackTransparent <> Null) Then
		If Not IsBool($bBackTransparent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oCell.BackTransparent = $bBackTransparent
		$iError = ($oCell.BackTransparent() = $bBackTransparent) ? $iError : BitOR($iError, 2) ; Error setting BackTransparent.
	EndIf

	Return ($iError = 0) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0)
EndFunc   ;==>_LOWriter_CellBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBorderColor
; Description ...: Set the Cell or Cell Range Border Line Color. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_CellBorderColor(Byref $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from any Table Cell Object creation or retrieval functions.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Color of the Cell in Long Color code format.
;				   +						One of the predefined constants listed below can be used, or a custom value.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Color of the Cell in Long Color code format.
;				   +						One of the predefined constants listed below can be used, or a custom value.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Color of the Cell in Long Color code format.
;				   +						One of the predefined constants listed below can be used, or a custom value.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Color of the Cell in Long Color code format.
;				   +						One of the predefined constants listed below can be used, or a custom value.
; Internal Remark: Error values for Initialization and Processing are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell Variable not Object type variable.
;				   @Error 1 @Extended 2 Return 0 = $iTop not an integer, or set to less than 0 or higher than 16,777,215.
;				   @Error 1 @Extended 3 Return 0 = $iBottom not an integer, or set to less than 0 or higher than 16,777,215.
;				   @Error 1 @Extended 4 Return 0 = $iLeft not an integer, or set to less than 0 or higher than 16,777,215.
;				   @Error 1 @Extended 5 Return 0 = $iRight not an integer, or set to less than 0 or higher than 16,777,215.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Color when Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Color when Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border Color when Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border Color when Border width not set.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Color Constants: $LOW_COLOR_BLACK(0),
;					$LOW_COLOR_WHITE(16777215),
;					$LOW_COLOR_LGRAY(11711154),
;					$LOW_COLOR_GRAY(8421504),
;					$LOW_COLOR_DKGRAY(3355443),
;					$LOW_COLOR_YELLOW(16776960),
;					$LOW_COLOR_GOLD(16760576),
;					$LOW_COLOR_ORANGE(16744448),
;					$LOW_COLOR_BRICK(16728064),
;					$LOW_COLOR_RED(16711680),
;					$LOW_COLOR_MAGENTA(12517441),
;					$LOW_COLOR_PURPLE(8388736),
;					$LOW_COLOR_INDIGO(5582989),
;					$LOW_COLOR_BLUE(2777241),
;					$LOW_COLOR_TEAL(1410150),
;					$LOW_COLOR_GREEN(43315),
;					$LOW_COLOR_LIME(8508442),
;					$LOW_COLOR_BROWN(9127187).
; Related .......: _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong, _LOWriter_CellBorderWidth,
;					_LOWriter_CellBorderStyle, _LOWriter_CellBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellBorderColor(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oCell, False, False, True, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)

EndFunc   ;==>_LOWriter_CellBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBorderPadding
; Description ...: Set or retrieve the Border Padding (spacing between the Cell text and border) settings.
; Syntax ........: _LOWriter_CellBorderPadding(Byref $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from any Table Cell Object creation or retrieval functions.
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Distance between the Border and Cell text in Micrometers(uM).
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Distance between the Border and Cell text in Micrometers(uM).
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Distance between the Border and Cell text in Micrometers(uM).
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Distance between the Border and Cell text in Micrometers(uM).
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell parameter not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iTop not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iBottom not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $Left not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iRight not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iTop border distance
;				   |								2 = Error setting $iBottom border distance
;				   |								4 = Error setting $iLeft border distance
;				   |								8 = Error setting $iRight border distance
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_CellBorderColor,
;					_LOWriter_CellBorderStyle, _LOWriter_CellBorderWidth
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellBorderPadding(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[4]

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		__LOWriter_ArrayFill($aiBPadding, $oCell.TopBorderDistance(), $oCell.BottomBorderDistance(), $oCell.LeftBorderDistance(), $oCell.RightBorderDistance())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iTop <> Null) Then
		If Not __LOWriter_IntIsBetween($iTop, 0, $iTop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oCell.TopBorderDistance = $iTop
		$iError = (__LOWriter_IntIsBetween($oCell.TopBorderDistance(), $iTop - 1, $iTop + 1)) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iBottom <> Null) Then
		If Not __LOWriter_IntIsBetween($iBottom, 0, $iBottom) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oCell.BottomBorderDistance = $iBottom
		$iError = (__LOWriter_IntIsBetween($oCell.BottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iLeft <> Null) Then
		If Not __LOWriter_IntIsBetween($iLeft, 0, $iLeft) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oCell.LeftBorderDistance = $iLeft
		$iError = (__LOWriter_IntIsBetween($oCell.LeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iRight <> Null) Then
		If Not __LOWriter_IntIsBetween($iRight, 0, $iRight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oCell.RightBorderDistance = $iRight
		$iError = (__LOWriter_IntIsBetween($oCell.RightBorderDistance(), $iRight - 1, $iRight + 1)) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CellBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBorderStyle
; Description ...: Set or Retrieve the Cell or Cell Range Border Line style. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_CellBorderStyle(Byref $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from any Table Cell Object creation or retrieval functions.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Style of theCell using one of the line style constants, See below for list.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Style of the Cell using one of the line style constants, See below for list.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Style of the Cell using one of the line style constants, See below for list.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Style of the Cell using one of the line style constants, See below for list.
; Internal Remark: Error values for Initialization and Processing are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell Variable not Object type variable.
;				   @Error 1 @Extended 2 Return 0 = $iTop not an integer, or set to higher than 17 and not equal to 0x7FFF, or $iTop is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 3 Return 0 = $iBottom not an integer, or set to higher than 17 and not equal to 0x7FFF, or $iBottom is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iLeft not an integer, or set to higher than 17 and not equal to 0x7FFF, or $iLeft is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iRight not an integer, or set to higher than 17 and not equal to 0x7FFF, or $iRight is set to less than 0 or not set to Null.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style when Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style when Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style when Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style when Border width not set.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Style Constants: $LOW_BORDERSTYLE_NONE(0x7FFF) No border line,
;					$LOW_BORDERSTYLE_SOLID(0) Solid border line,
;					$LOW_BORDERSTYLE_DOTTED(1) Dotted border line,
;					$LOW_BORDERSTYLE_DASHED(2) Dashed border line,
;					$LOW_BORDERSTYLE_DOUBLE(3) Double border line,
;					$LOW_BORDERSTYLE_THINTHICK_SMALLGAP(4) Double border line with a thin line outside and a thick line inside separated by a small gap,
;					$LOW_BORDERSTYLE_THINTHICK_MEDIUMGAP(5) Double border line with a thin line outside and a thick line inside separated by a medium gap,
;					$LOW_BORDERSTYLE_THINTHICK_LARGEGAP(6) Double border line with a thin line outside and a thick line inside separated by a large gap,
;					$LOW_BORDERSTYLE_THICKTHIN_SMALLGAP(7) Double border line with a thick line outside and a thin line inside separated by a small gap,
;					$LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP(8) Double border line with a thick line outside and a thin line inside separated by a medium gap,
;					$LOW_BORDERSTYLE_THICKTHIN_LARGEGAP(9) Double border line with a thick line outside and a thin line inside separated by a large gap,
;					$LOW_BORDERSTYLE_EMBOSSED(10) 3D embossed border line,
;					$LOW_BORDERSTYLE_ENGRAVED(11) 3D engraved border line,
;					$LOW_BORDERSTYLE_OUTSET(12) Outset border line,
;					$LOW_BORDERSTYLE_INSET(13) Inset border line,
;					$LOW_BORDERSTYLE_FINE_DASHED(14) Finely dashed border line,
;					$LOW_BORDERSTYLE_DOUBLE_THIN(15) Double border line consisting of two fixed thin lines separated by a variable gap,
;					$LOW_BORDERSTYLE_DASH_DOT(16) Line consisting of a repetition of one dash and one dot,
;					$LOW_BORDERSTYLE_DASH_DOT_DOT(17) Line consisting of a repetition of one dash and 2 dots.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_CellBorderWidth, _LOWriter_CellBorderColor, _LOWriter_CellBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellBorderStyle(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oCell, False, True, False, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CellBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBorderWidth
; Description ...: Set or Retrieve the Cell or Cell Range Border Line Width. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_CellBorderWidth(Byref $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from any Table Cell
;				   +						Object creation or retrieval functions.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line width of the
;				   +						Cell in MicroMeters. One of the predefined constants listed below can be used.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Width of the
;				   +							Cell in MicroMeters. One of the predefined constants listed below can be used.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line width of the
;				   +							Cell in MicroMeters. One of the predefined constants listed below can be used.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Width of the
;				   +							Cell in MicroMeters. One of the predefined constants listed below can be used.
; Internal Remark: Error values for Initialization and Processing are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell Variable not Object type variable.
;				   @Error 1 @Extended 2 Return 0 = $iTop not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 3 Return 0 = $iBottom not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iLeft not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iRight not an integer, or set to less than 0 or not set to Null.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set them to 0
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Width Constants: $LOW_BORDERWIDTH_HAIRLINE(2),
;					$LOW_BORDERWIDTH_VERY_THIN(18),
;					$LOW_BORDERWIDTH_THIN(26),
;					$LOW_BORDERWIDTH_MEDIUM(53),
;					$LOW_BORDERWIDTH_THICK(79),
;					$LOW_BORDERWIDTH_EXTRA_THICK(159)
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_CellBorderStyle,
;					_LOWriter_CellBorderColor, _LOWriter_CellBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellBorderWidth(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, 0, $iTop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, 0, $iBottom) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, 0, $iLeft) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, 0, $iRight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_Border($oCell, True, False, False, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_CellBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellCreateTextCursor
; Description ...: Create a Text Cursor in a particular cell for inserting text etc.
; Syntax ........: _LOWriter_CellCreateTextCursor(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object creation or retrieval functions.
; Return values .: Success: An Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. A Text Cursor Object located in the specified Cell.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellCreateTextCursor(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; Can only create a Text Cursor for individual cells.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.Text.createTextCursor())
EndFunc   ;==>_LOWriter_CellCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellFormula
; Description ...: Set or retrieve a formula for a cell.
; Syntax ........: _LOWriter_CellFormula(Byref $oCell[, $sFormula = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object creation or retrieval functions.
;                  $sFormula            - [optional] a string value. Default is Null. The Formula to set the Cell to.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFormula not a String and not set to Null keyword.
;				   @Error 1 @Extended 3 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Formula was successfully set.
;				   @Error 0 @Extended 1 Return String = Success. Current formula is returned in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Formula can only be set for an individual cell, not a range.
;					Setting the formula will overwrite any existing data in the cell.
; 					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					To retrieve the total of a formula, use _LOWriter_CellValue.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellFormula(ByRef $oCell, $sFormula = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFormula) And Not ($sFormula = Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; Can only set/get formula value for individual cells.
	If ($sFormula = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oCell.getFormula())

	$oCell.setFormula($sFormula)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CellFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellGetDataType
; Description ...: Get the Data type of a specific cell, see remarks.
; Syntax ........: _LOWriter_CellGetDataType(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object creation or retrieval functions.
; Return values .: Success: A Number.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return Number = Success. The Data Type in Number format, see constants below.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Returns the data type as one of the below constants, Note: If the data was entered by the keyboard, it is generally recognized as a string regardless of the data contained.
; Data Type Constants: $LOW_CELL_TYPE_EMPTY(0), cell is empty.
;						$LOW_CELL_TYPE_VALUE(1), cell contains a value.
;						$LOW_CELL_TYPE_TEXT(2), cell contains text.
;						$LOW_CELL_TYPE_FORMULA(3), cell contains a formula.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellGetDataType(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; Can only get Data Type for individual cells

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.getType())
EndFunc   ;==>_LOWriter_CellGetDataType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellGetError
; Description ...: Get the formula error Value.
; Syntax ........: _LOWriter_CellGetError(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object creation or retrieval functions.
; Return values .: Success: A Number.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return Number = Success. The Cell formula error code in Number format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Integer error value. If the cell is not a formula, the error value is zero.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellGetError(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; Can only get Error for individual cells.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.getError())

EndFunc   ;==>_LOWriter_CellGetError

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellGetName
; Description ...: Retrieve the current Cell's name.
; Syntax ........: _LOWriter_CellGetName(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object creation or retrieval functions.
; Return values .: Success: A String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. The Cell name in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellGetName(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; Can only get Cell Name for individual cells.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.CellName())
EndFunc   ;==>_LOWriter_CellGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellProtect
; Description ...: Write-Protect a Cell
; Syntax ........: _LOWriter_CellProtect(Byref $oCell[, $bProtect = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object creation or retrieval functions.
;                  $bProtect            - [optional] a boolean value. Default is Null. True = Protected from Writing, False = Unprotected. See remarks.
; Return values .: Success: 1 Or Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range. Can only set Write-Protect on individual cells.
;				   @Error 1 @Extended 3 Return 0 = $bProtect not a Boolean or not Null keyword.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Failed to set Write-Protect property.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully set Cell Protect setting.
;				   @Error 0 @Extended 0 Return Boolean = Success. $bProtect is set to Null, return will be the current setting of write-protection for the cell, a Boolean value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling $bProtect with Null keyword returns the current WriteProtection setting of the cell. (True or
;					False)
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellProtect(ByRef $oCell, $bProtect = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; Can only set individual cell protect property.
	If ($bProtect = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.IsProtected())
	If Not IsBool($bProtect) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oCell.IsProtected = $bProtect
	Return ($oCell.IsProtected() = $bProtect) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)

EndFunc   ;==>_LOWriter_CellProtect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellString
; Description ...: Set or retrieve the current string for a cell.
; Syntax ........: _LOWriter_CellString(Byref $oCell[, $sString = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object creation or retrieval functions.
;                  $sString             - [optional] a string value. Default is Null. The String of text to set the cell to.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sString not a String and not set to Null keyword.
;				   @Error 1 @Extended 3 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. String was successfully set.
;				   @Error 0 @Extended 1 Return String = Success. Current String is returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: String can only be set for an individual cell, not a range.
;					Setting the String will overwrite any existing data in the cell.
; 					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellString(ByRef $oCell, $sString = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sString) And Not ($sString = Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; Can only set/get a String for individual cells.

	If ($sString = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oCell.getString())

	$oCell.setString($sString)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CellString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellValue
; Description ...: Set or retrieve a Numerical value to a Cell
; Syntax ........: _LOWriter_CellValue(Byref $oCell[, $nValue = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object creation or retrieval functions.
;                  $nValue              - [optional] a general number value. Default is Null. The value to set the cell to.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $nValue not a Number and not set to Null keyword.
;				   @Error 1 @Extended 3 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Value was successfully set.
;				   @Error 0 @Extended 1 Return String = Success. Current Value is returned in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Value can only be set for an individual cell, not a range.
;					Setting the Value will overwrite any existing data in the cell.
;					For a value cell the value is returned, for a string cell zero is returned and for a formula cell the result value of a formula is returned.
; 					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellValue(ByRef $oCell, $nValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsNumber($nValue) And Not ($nValue = Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; Can only set/get individual cell values.

	If ($nValue = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oCell.getValue())

	$oCell.setValue($nValue)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CellValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellVertOrient
; Description ...: Set the Vertical Orientation of the Cell or Cell Range contents.
; Syntax ........: _LOWriter_CellVertOrient(Byref $oCell[, $iVertOrient = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from any Table Cell
;				   +						Object creation or retrieval functions.
;                  $iVertOrient         - [optional]  an integer value. Default is Null. A Vertical Orientation constant. See
;				   +						Constants Below.
; Return values .: Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iVertOrient not an integer, or less than 0 or greater than 3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Failed to set Cell Vertical Orientation property.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1  = Success. Succesfully set Vertical Orientation.
;				   @Error 0 @Extended 0 Return Integer = Success. $iVertOrient is set to Null, returning current Cell Vertical orientation, see constants below.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only the Vertical Orientation Constants listed below are accepted. If $iVertOrient is set to Null the current setting is returned.
; Vertical Orientation Constants: $LOW_ORIENT_VERT_NONE(0),
;									$LOW_ORIENT_VERT_TOP(1),
;									$LOW_ORIENT_VERT_CENTER(2),
;									$LOW_ORIENT_VERT_BOTTOM(3)
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellVertOrient(ByRef $oCell, $iVertOrient = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	; 3 = Vert Orient Bottom, 1 = Vert orient Top

	If ($iVertOrient = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.VertOrient())
	If Not __LOWriter_IntIsBetween($iVertOrient, $LOW_ORIENT_VERT_NONE, $LOW_ORIENT_VERT_BOTTOM) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oCell.VertOrient = $iVertOrient

	Return ($oCell.VertOrient() = $iVertOrient) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)
EndFunc   ;==>_LOWriter_CellVertOrient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteDelete
; Description ...: Delete a Endnote.
; Syntax ........: _LOWriter_EndnoteDelete(Byref $oEndNote)
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous Endnote insert, or retrieval function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oEndNote not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Endnote successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteDelete(ByRef $oEndNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oEndNote.dispose()
	$oEndNote = Null

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_EndnoteDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteGetAnchor
; Description ...: Create a Text Cursor at the Endnote Anchor position.
; Syntax ........: _LOWriter_EndnoteGetAnchor(Byref $oEndNote)
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous Endnote insert, or retrieval function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oEndNote not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully returned the Endnote Anchor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Anchor cursor returned is just a Text Cursor placed at the anchor's position.
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert, _LOWriter_CursorMove, _LOWriter_DocGetString, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteGetAnchor(ByRef $oEndNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnchor

	If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oAnchor = $oEndNote.Anchor.Text.createTextCursorByRange($oEndNote.Anchor())
	If Not IsObj($oAnchor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oAnchor)
EndFunc   ;==>_LOWriter_EndnoteGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteGetTextCursor
; Description ...: Create a Text Cursor in a Endnote to modify the text therein.
; Syntax ........: _LOWriter_EndnoteGetTextCursor(Byref $oEndNote)
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous Endnote insert, or retrieval function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oEndNote not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Cursor Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully retrieved the Endnote Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteGetTextCursor(ByRef $oEndNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oTextCursor = $oEndNote.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOWriter_EndnoteGetTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteInsert
; Description ...: Insert a Endnote into a Document.
; Syntax ........: _LOWriter_EndnoteInsert(Byref $oDoc, Byref $oCursor, $bOverwrite[, $sLabel = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten.
;				   +								If False, content will be inserted to the left of any selection.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the Endnote.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $oCursor is a Table cursor type, not supported.
;				   @Error 1 @Extended 5 Return 0 = $oCursor currently located in a Frame, Footnote, Endnote, or Header/ Footer cannot insert a Endnote in those data types.
;				   @Error 1 @Extended 6 Return 0 = $oCursor located in unknown data type.
;				   @Error 1 @Extended 7 Return 0 = $sLabel not a string.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 =  Error creating "com.sun.star.text.Endnote" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully inserted a new Endnote, returning Endnote Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Endnote cannot be inserted into a Frame, a Footnote, a Endnote, or the Header/ Footer.
; Related .......: _LOWriter_EndnoteDelete,  _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sLabel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oEndNote

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	Switch __LOWriter_Internal_CursorGetDataType($oDoc, $oCursor)

		Case $LOW_CURDATA_FRAME, $LOW_CURDATA_FOOTNOTE, $LOW_CURDATA_ENDNOTE, $LOW_CURDATA_HEADER_FOOTER
			Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0) ; Unsupported cursor type.
		Case $LOW_CURDATA_BODY_TEXT, $LOW_CURDATA_CELL
			$oEndNote = $oDoc.createInstance("com.sun.star.text.Endnote")
			If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

		Case Else
			Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0) ; Unknown Cursor type.
	EndSwitch

	If ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oEndNote.Label = $sLabel
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oEndNote, $bOverwrite)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oEndNote)
EndFunc   ;==>_LOWriter_EndnoteInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteModifyAnchor
; Description ...: Modify a Specific Endnote's settings.
; Syntax ........: _LOWriter_EndnoteModifyAnchor(Byref $oEndNote[, $sLabel = Null])
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous Endnote insert, or retrieval function.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the Endnote. Set to "" for automatic numbering.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oEndNote not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sLabel not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = $sLabel was not set successfully.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Endnote settings were successfully modified.
;				   @Error 0 @Extended 1 Return String = Success. $sLabel set to Null, current Endnote Label returned.
;				   @Error 0 @Extended 2 Return String = Success. $sLabel set to Null, current Endnote AutoNumbering number returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteModifyAnchor(ByRef $oEndNote, $sLabel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($sLabel = Null) Then
		; If Label is blank, return the AutoNumbering Number.
		If ($oEndNote.Label() = "") Then Return SetError($__LOW_STATUS_SUCCESS, 2, $oEndNote.Anchor.String())

		; Else return the Label.
		Return SetError($__LOW_STATUS_SUCCESS, 1, $oEndNote.Label())

	EndIf

	If Not IsString($sLabel) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oEndNote.Label = $sLabel
	If ($oEndNote.Label() <> $sLabel) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_EndnoteModifyAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteSettingsAutoNumber
; Description ...: Set or Retrieve Endnote Autonumbering settings.
; Syntax ........: _LOWriter_EndnoteSettingsAutoNumber(Byref $oDoc[, $iNumFormat = Null[, $iStartAt = Null[, $sBefore = Null[, $sAfter = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iNumFormat          - [optional] an integer value. Default is Null. The numbering format to use for Endnote numbering. See Constants.
;                  $iStartAt            - [optional] an integer value. Default is Null. The Number to begin Endnote counting from, Min. 1, Max 9999.
;                  $sBefore             - [optional] a string value. Default is Null. The text to display before a Endnote number in the note text.
;                  $sAfter              - [optional] a string value. Default is Null. The text to display after a Endnote number in the note text.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iNumFormat not an Integer, or Less than 0 or greater than 71. See Constants.
;				   @Error 1 @Extended 3 Return 0 = $iStartAt not an integer, less than 1 or greater than 9999.
;				   @Error 1 @Extended 4 Return 0 = $sBefore not a String.
;				   @Error 1 @Extended 5 Return 0 = $sAfter not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iNumFormat
;				   |								2 = Error setting $iStartAt
;				   |								4 = Error setting $sBefore
;				   |								8 = Error setting $sAfter
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Numbering Format Constants: $LOW_NUM_STYLE_CHARS_UPPER_LETTER(0), Numbering is put in upper case letters. ("A, B, C, D)
;	$LOW_NUM_STYLE_CHARS_LOWER_LETTER(1), Numbering is in lower case letters. (a, b, c, d)
;	$LOW_NUM_STYLE_ROMAN_UPPER(2), Numbering is in Roman numbers with upper case letters. (I, II, III)
;	$LOW_NUM_STYLE_ROMAN_LOWER(3), Numbering is in Roman numbers with lower case letters. (i, ii, iii)
;	$LOW_NUM_STYLE_ARABIC(4), Numbering is in Arabic numbers. (1, 2, 3, 4)
;	$LOW_NUM_STYLE_NUMBER_NONE(5), Numbering is invisible.
;	$LOW_NUM_STYLE_CHAR_SPECIAL(6), Use a character from a specified font.
;	$LOW_NUM_STYLE_PAGE_DESCRIPTOR(7), Numbering is specified in the page style.
;	$LOW_NUM_STYLE_BITMAP(8), Numbering is displayed as a bitmap graphic.
;	$LOW_NUM_STYLE_CHARS_UPPER_LETTER_N(9), Numbering is put in upper case letters. (A, B, Y, Z, AA, BB)
;	$LOW_NUM_STYLE_CHARS_LOWER_LETTER_N(10), Numbering is put in lower case letters. (a, b, y, z, aa, bb)
;	$LOW_NUM_STYLE_TRANSLITERATION(11), A transliteration module will be used to produce numbers in Chinese, Japanese, etc.
;	$LOW_NUM_STYLE_NATIVE_NUMBERING(12), The NativeNumberSupplier service will be called to produce numbers in native languages.
;	$LOW_NUM_STYLE_FULLWIDTH_ARABIC(13), Numbering for full width Arabic number.
;	$LOW_NUM_STYLE_CIRCLE_NUMBER(14), 	Bullet for Circle Number.
;	$LOW_NUM_STYLE_NUMBER_LOWER_ZH(15), Numbering for Chinese lower case number.
;	$LOW_NUM_STYLE_NUMBER_UPPER_ZH(16), Numbering for Chinese upper case number.
;	$LOW_NUM_STYLE_NUMBER_UPPER_ZH_TW(17), Numbering for Traditional Chinese upper case number.
;	$LOW_NUM_STYLE_TIAN_GAN_ZH(18), Bullet for Chinese Tian Gan.
;	$LOW_NUM_STYLE_DI_ZI_ZH(19), Bullet for Chinese Di Zi.
;	$LOW_NUM_STYLE_NUMBER_TRADITIONAL_JA(20), Numbering for Japanese traditional number.
;	$LOW_NUM_STYLE_AIU_FULLWIDTH_JA(21), Bullet for Japanese AIU fullwidth.
;	$LOW_NUM_STYLE_AIU_HALFWIDTH_JA(22), Bullet for Japanese AIU halfwidth.
;	$LOW_NUM_STYLE_IROHA_FULLWIDTH_JA(23), Bullet for Japanese IROHA fullwidth.
;	$LOW_NUM_STYLE_IROHA_HALFWIDTH_JA(24), Bullet for Japanese IROHA halfwidth.
;	$LOW_NUM_STYLE_NUMBER_UPPER_KO(25), Numbering for Korean upper case number.
;	$LOW_NUM_STYLE_NUMBER_HANGUL_KO(26), Numbering for Korean Hangul number.
;	$LOW_NUM_STYLE_HANGUL_JAMO_KO(27), Bullet for Korean Hangul Jamo.
;	$LOW_NUM_STYLE_HANGUL_SYLLABLE_KO(28), Bullet for Korean Hangul Syllable.
;	$LOW_NUM_STYLE_HANGUL_CIRCLED_JAMO_KO(29), Bullet for Korean Hangul Circled Jamo.
;	$LOW_NUM_STYLE_HANGUL_CIRCLED_SYLLABLE_KO(30), Bullet for Korean Hangul Circled Syllable.
;	$LOW_NUM_STYLE_CHARS_ARABIC(31), Numbering in Arabic alphabet letters.
;	$LOW_NUM_STYLE_CHARS_THAI(32), Numbering in Thai alphabet letters.
;	$LOW_NUM_STYLE_CHARS_HEBREW(33), Numbering in Hebrew alphabet letters.
;	$LOW_NUM_STYLE_CHARS_NEPALI(34), Numbering in Nepali alphabet letters.
;	$LOW_NUM_STYLE_CHARS_KHMER(35), Numbering in Khmer alphabet letters.
;	$LOW_NUM_STYLE_CHARS_LAO(36), Numbering in Lao alphabet letters.
;	$LOW_NUM_STYLE_CHARS_TIBETAN(37), Numbering in Tibetan/Dzongkha alphabet letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_BG(38), Numbering in Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_BG(39), Numbering in Cyrillic alphabet lower case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_BG(40), Numbering in Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_BG(41), Numbering in Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_RU(42), Numbering in Russian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_RU(43), Numbering in Russian Cyrillic alphabet lower case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_RU(44), Numbering in Russian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_RU(45), Numbering in Russian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_PERSIAN(46), Numbering in Persian alphabet letters.
;	$LOW_NUM_STYLE_CHARS_MYANMAR(47), Numbering in Myanmar alphabet letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_SR(48), Numbering in Serbian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_SR(49), Numbering in Russian Serbian alphabet lower case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_SR(50), Numbering in Serbian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_SR(51), Numbering in Serbian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_GREEK_UPPER_LETTER(52), Numbering in Greek alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_GREEK_LOWER_LETTER(53), Numbering in Greek alphabet lower case letters.
;	$LOW_NUM_STYLE_CHARS_ARABIC_ABJAD(54), Numbering in Arabic alphabet using abjad sequence.
;	$LOW_NUM_STYLE_CHARS_PERSIAN_WORD(55), Numbering in Persian words.
;	$LOW_NUM_STYLE_NUMBER_HEBREW(56), Numbering in Hebrew numerals.
;	$LOW_NUM_STYLE_NUMBER_ARABIC_INDIC(57), Numbering in Arabic-Indic numerals.
;	$LOW_NUM_STYLE_NUMBER_EAST_ARABIC_INDIC(58), Numbering in East Arabic-Indic numerals.
;	$LOW_NUM_STYLE_NUMBER_INDIC_DEVANAGARI(59), Numbering in Indic Devanagari numerals.
;	$LOW_NUM_STYLE_TEXT_NUMBER(60), Numbering in ordinal numbers of the language of the text node. (1st, 2nd, 3rd)
;	$LOW_NUM_STYLE_TEXT_CARDINAL(61), Numbering in cardinal numbers of the language of the text node. (One, Two)
;	$LOW_NUM_STYLE_TEXT_ORDINAL(62), Numbering in ordinal numbers of the language of the text node. (First, Second)
;	$LOW_NUM_STYLE_SYMBOL_CHICAGO(63), Footnoting symbols according the University of Chicago style.
;	$LOW_NUM_STYLE_ARABIC_ZERO(64), Numbering is in Arabic numbers, padded with zero to have a length of at least two. (01, 02)
;	$LOW_NUM_STYLE_ARABIC_ZERO3(65), Numbering is in Arabic numbers, padded with zero to have a length of at least three.
;	$LOW_NUM_STYLE_ARABIC_ZERO4(66), Numbering is in Arabic numbers, padded with zero to have a length of at least four.
;	$LOW_NUM_STYLE_ARABIC_ZERO5(67), Numbering is in Arabic numbers, padded with zero to have a length of at least five.
;	$LOW_NUM_STYLE_SZEKELY_ROVAS(68), Numbering is in Szekely rovas (Old Hungarian) numerals.
;	$LOW_NUM_STYLE_NUMBER_DIGITAL_KO(69), Numbering is in Korean Digital number.
;	$LOW_NUM_STYLE_NUMBER_DIGITAL2_KO(70), Numbering is in Korean Digital Number, reserved "koreanDigital2".
;	$LOW_NUM_STYLE_NUMBER_LEGAL_KO(71), Numbering is in Korean Legal Number, reserved "koreanLegal".
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteSettingsAutoNumber(ByRef $oDoc, $iNumFormat = Null, $iStartAt = Null, $sBefore = Null, $sAfter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avENSettings[4]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iNumFormat, $iStartAt, $sBefore, $sAfter) Then
		__LOWriter_ArrayFill($avENSettings, $oDoc.EndnoteSettings.NumberingType(), ($oDoc.EndnoteSettings.StartAt() + 1), _
				$oDoc.EndnoteSettings.Prefix(), $oDoc.EndnoteSettings.Suffix())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avENSettings)
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDoc.EndnoteSettings.NumberingType = $iNumFormat
		$iError = ($oDoc.EndnoteSettings.NumberingType() = $iNumFormat) ? $iError : BitOR($iError, 1)
	EndIf

	; 0 Based -- Minus 1
	If ($iStartAt <> Null) Then
		If Not __LOWriter_IntIsBetween($iStartAt, 1, 9999) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDoc.EndnoteSettings.StartAt = ($iStartAt - 1)
		$iError = ($oDoc.EndnoteSettings.StartAt() = ($iStartAt - 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sBefore <> Null) Then
		If Not IsString($sBefore) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oDoc.EndnoteSettings.Prefix = $sBefore
		$iError = ($oDoc.EndnoteSettings.Prefix() = $sBefore) ? $iError : BitOR($iError, 4)
	EndIf

	If ($sAfter <> Null) Then
		If Not IsString($sAfter) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDoc.EndnoteSettings.Suffix = $sAfter
		$iError = ($oDoc.EndnoteSettings.Suffix() = $sAfter) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_EndnoteSettingsAutoNumber

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteSettingsStyles
; Description ...: Set or Retrieve Document Endnote Style settings.
; Syntax ........: _LOWriter_EndnoteSettingsStyles(Byref $oDoc[, $sParagraph = Null[, $sPage = Null[, $sTextArea = Null[, $sEndnoteArea = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sParagraph          - [optional] a string value. Default is Null. The Endnote Text Paragraph Style.
;                  $sPage               - [optional] a string value. Default is Null. The Page Style to use for the Endnote pages.
;                  $sTextArea           - [optional] a string value. Default is Null. The Character Style to use for the Endnote anchor in the document text.
;                  $sEndnoteArea        - [optional] a string value. Default is Null. The Character Style to use for the Endnote number in the Endnote text.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sParagraph not a String.
;				   @Error 1 @Extended 3 Return 0 = Paragraph Style referenced in $sParagraph not found in Document.
;				   @Error 1 @Extended 4 Return 0 = $sPage not a String.
;				   @Error 1 @Extended 5 Return 0 = Page Style referenced in $sPage not found in Document.
;				   @Error 1 @Extended 6 Return 0 = $sTextArea not a String.
;				   @Error 1 @Extended 7 Return 0 = Character Style referenced in $sTextArea not found in Document.
;				   @Error 1 @Extended 8 Return 0 = $sEndnoteArea not a String.
;				   @Error 1 @Extended 9 Return 0 = Character Style referenced in $sEndnoteArea not found in Document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sParagraph
;				   |								2 = Error setting $sPage
;				   |								4 = Error setting $sTextArea
;				   |								8 = Error setting $sEndnoteArea
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStylesGetNames, _LOWriter_CharStylesGetNames, _LOWriter_PageStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteSettingsStyles(ByRef $oDoc, $sParagraph = Null, $sPage = Null, $sTextArea = Null, $sEndnoteArea = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asENSettings[4]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sParagraph, $sPage, $sTextArea, $sEndnoteArea) Then
		__LOWriter_ArrayFill($asENSettings, __LOWriter_ParStyleNameToggle($oDoc.EndnoteSettings.ParaStyleName(), True), _
				__LOWriter_PageStyleNameToggle($oDoc.EndnoteSettings.PageStyleName(), True), _
				__LOWriter_CharStyleNameToggle($oDoc.EndnoteSettings.AnchorCharStyleName(), True), _
				__LOWriter_CharStyleNameToggle($oDoc.EndnoteSettings.CharStyleName(), True))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $asENSettings)
	EndIf

	If ($sParagraph <> Null) Then
		If Not IsString($sParagraph) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		If Not _LOWriter_ParStyleExists($oDoc, $sParagraph) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$sParagraph = __LOWriter_ParStyleNameToggle($sParagraph)
		$oDoc.EndnoteSettings.ParaStyleName = $sParagraph
		$iError = ($oDoc.EndnoteSettings.ParaStyleName() = $sParagraph) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sPage <> Null) Then
		If Not IsString($sPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		If Not _LOWriter_PageStyleExists($oDoc, $sPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$sPage = __LOWriter_PageStyleNameToggle($sPage)
		$oDoc.EndnoteSettings.PageStyleName = $sPage
		$iError = ($oDoc.EndnoteSettings.PageStyleName() = $sPage) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sTextArea <> Null) Then
		If Not IsString($sTextArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_CharStyleExists($oDoc, $sTextArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$sTextArea = __LOWriter_CharStyleNameToggle($sTextArea)
		$oDoc.EndnoteSettings.AnchorCharStyleName = $sTextArea
		$iError = ($oDoc.EndnoteSettings.AnchorCharStyleName() = $sTextArea) ? $iError : BitOR($iError, 4)
	EndIf

	If ($sEndnoteArea <> Null) Then
		If Not IsString($sEndnoteArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		If Not _LOWriter_CharStyleExists($oDoc, $sEndnoteArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$sEndnoteArea = __LOWriter_CharStyleNameToggle($sEndnoteArea)
		$oDoc.EndnoteSettings.CharStyleName = $sEndnoteArea
		$iError = ($oDoc.EndnoteSettings.CharStyleName() = $sEndnoteArea) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_EndnoteSettingsStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnotesGetList
; Description ...: Retrieve an array of Endnote Objects contained in a Document.
; Syntax ........: _LOWriter_EndnotesGetList(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Endnotes Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully searched for Endnotes, none contained in document.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for Endnotes, Returning Array of Endnote Objects. @Extended set to number found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_EndnoteDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnotesGetList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oEndNotes
	Local $aoEndnotes[0]
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oEndNotes = $oDoc.getEndnotes()
	If Not IsObj($oEndNotes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$iCount = $oEndNotes.getCount()

	If ($iCount > 0) Then
		ReDim $aoEndnotes[$iCount]

		For $i = 0 To $iCount - 1
			$aoEndnotes[$i] = $oEndNotes.getByIndex($i)

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return ($iCount > 0) ? SetError($__LOW_STATUS_SUCCESS, $iCount, $aoEndnotes) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_EndnotesGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteDelete
; Description ...: Delete a Footnote.
; Syntax ........: _LOWriter_FootnoteDelete(Byref $oFootNote)
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous Footnote insert, or retrieval function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oFootNote not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Footnote successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FootnoteInsert, _LOWriter_FootnotesGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteDelete(ByRef $oFootNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oFootNote.dispose()
	$oFootNote = Null

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnoteDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteGetAnchor
; Description ...: Create a Text Cursor at the Footnote Anchor position.
; Syntax ........: _LOWriter_FootnoteGetAnchor(Byref $oFootNote)
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous Footnote insert, or retrieval function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oFootNote not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully returned the Footnote Anchor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FootnotesGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteGetAnchor(ByRef $oFootNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnchor

	If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oAnchor = $oFootNote.Anchor.Text.createTextCursorByRange($oFootNote.Anchor())
	If Not IsObj($oAnchor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oAnchor)
EndFunc   ;==>_LOWriter_FootnoteGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteGetTextCursor
; Description ...: Create a Text Cursor in a Footnote to modify the text therein.
; Syntax ........: _LOWriter_FootnoteGetTextCursor(Byref $oFootNote)
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous Footnote insert, or retrieval function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oFootNote not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Cursor Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully retrieved the footnote Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_CursorMove, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteGetTextCursor(ByRef $oFootNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oTextCursor = $oFootNote.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOWriter_FootnoteGetTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteInsert
; Description ...: Insert a Footnote into a Document.
; Syntax ........: _LOWriter_FootnoteInsert(Byref $oDoc, Byref $oCursor, $bOverwrite[, $sLabel = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten.
;				   +									If False, content will be inserted to the left of any selection.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the footnote.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $oCursor is a Table cursor type, not supported.
;				   @Error 1 @Extended 5 Return 0 = $oCursor currently located in a Frame, Footnote, Endnote, or Header/Footer,
;				   +									cannot insert a Footnote in those data types.
;				   @Error 1 @Extended 6 Return 0 = $oCursor located in unknown data type.
;				   @Error 1 @Extended 7 Return 0 = $sLabel not a string.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 =  Error creating "com.sun.star.text.Footnote" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully inserted a new footnote, returning Footnote
;				   +									Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Footnote cannot be inserted into a Frame, a Footnote, a Endnote, or a Header/ Footer.
; Related .......: _LOWriter_FootnoteDelete, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sLabel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFootNote

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	Switch __LOWriter_Internal_CursorGetDataType($oDoc, $oCursor)

		Case $LOW_CURDATA_FRAME, $LOW_CURDATA_FOOTNOTE, $LOW_CURDATA_ENDNOTE, $LOW_CURDATA_HEADER_FOOTER
			Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0) ; Unsupported cursor type.
		Case $LOW_CURDATA_BODY_TEXT, $LOW_CURDATA_CELL
			$oFootNote = $oDoc.createInstance("com.sun.star.text.Footnote")
			If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

		Case Else
			Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0) ; Unknown Cursor type.
	EndSwitch

	If ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oFootNote.Label = $sLabel
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oFootNote, $bOverwrite)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oFootNote)
EndFunc   ;==>_LOWriter_FootnoteInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteModifyAnchor
; Description ...: Modify a Footnote's Anchor Character.
; Syntax ........: _LOWriter_FootnoteModifyAnchor(Byref $oFootNote[, $sLabel = Null])
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous Footnote insert, or retrieval function.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the Footnote. Set to "" for automatic numbering.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oFootNote not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sLabel not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Failed to set $sLabel.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Footnote settings were successfully modified.
;				   @Error 0 @Extended 1 Return String = Success. $sLabel set to Null, current Footnote Custom Label returned.
;				   @Error 0 @Extended 2 Return String = Success. $sLabel set to Null, current Footnote AutoNumbering number returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FootnoteInsert, _LOWriter_FootnotesGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteModifyAnchor(ByRef $oFootNote, $sLabel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($sLabel = Null) Then
		; If Label is blank, return the AutoNumbering Number.
		If ($oFootNote.Label() = "") Then Return SetError($__LOW_STATUS_SUCCESS, 2, $oFootNote.Anchor.String())

		; Else return the Label.
		Return SetError($__LOW_STATUS_SUCCESS, 1, $oFootNote.Label())
	EndIf

	If Not IsString($sLabel) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oFootNote.Label = $sLabel
	If ($oFootNote.Label() <> $sLabel) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnoteModifyAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteSettingsAutoNumber
; Description ...: Set or Retrieve Footnote Autonumbering settings.
; Syntax ........: _LOWriter_FootnoteSettingsAutoNumber(Byref $oDoc[, $iNumFormat = Null[, $iStartAt = Null[, $sBefore = Null[, $sAfter = Null[, $iCounting = Null[, $bEndOfDoc = Null]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iNumFormat            - [optional] an integer value. Default is Null. The numbering format to use for Footnote numbering. See Constants.
;                  $iStartAt            - [optional] an integer value. Default is Null. The Number to begin Footnote counting from, this is labeled "Counting" in the L.O. User Interface. Min. 1, Max 9999.
;                  $sBefore             - [optional] a string value. Default is Null. The text to display before a Footnote number in the note text.
;                  $sAfter              - [optional] a string value. Default is Null. The text to display after a Footnote number in the note text.
;                  $iCounting           - [optional] an integer value. Default is Null. The Counting type of the footnotes, such as per page etc., see constants below.
;                  $bEndOfDoc           - [optional] a boolean value. Default is Null. If True, Footnotes are placed at the end of the document, like Endnotes.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iNumFormat not an Integer, Less than 0 or greater than 71. See Constants.
;				   @Error 1 @Extended 3 Return 0 = $iStartAt not an integer, less than 1 or greater than 9999.
;				   @Error 1 @Extended 4 Return 0 = $sBefore not a String.
;				   @Error 1 @Extended 5 Return 0 = $sAfter not a String.
;				   @Error 1 @Extended 6 Return 0 = $iCounting not an Integer, less than 0 or greater than 2. See Constants.
;				   @Error 1 @Extended 7 Return 0 = $bEndOfDoc not a boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iNumFormat
;				   |								2 = Error setting $iStartAt
;				   |								4 = Error setting $sBefore
;				   |								8 = Error setting $sAfter
;				   |								16 = Error setting $iCounting
;				   |								32 = Error setting $bEndOfDoc
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Numbering Format Constants: $LOW_NUM_STYLE_CHARS_UPPER_LETTER(0), Numbering is put in upper case letters. ("A, B, C, D)
;	$LOW_NUM_STYLE_CHARS_LOWER_LETTER(1), Numbering is in lower case letters. (a, b, c, d)
;	$LOW_NUM_STYLE_ROMAN_UPPER(2), Numbering is in Roman numbers with upper case letters. (I, II, III)
;	$LOW_NUM_STYLE_ROMAN_LOWER(3), Numbering is in Roman numbers with lower case letters. (i, ii, iii)
;	$LOW_NUM_STYLE_ARABIC(4), Numbering is in Arabic numbers. (1, 2, 3, 4)
;	$LOW_NUM_STYLE_NUMBER_NONE(5), Numbering is invisible.
;	$LOW_NUM_STYLE_CHAR_SPECIAL(6), Use a character from a specified font.
;	$LOW_NUM_STYLE_PAGE_DESCRIPTOR(7), Numbering is specified in the page style.
;	$LOW_NUM_STYLE_BITMAP(8), Numbering is displayed as a bitmap graphic.
;	$LOW_NUM_STYLE_CHARS_UPPER_LETTER_N(9), Numbering is put in upper case letters. (A, B, Y, Z, AA, BB)
;	$LOW_NUM_STYLE_CHARS_LOWER_LETTER_N(10), Numbering is put in lower case letters. (a, b, y, z, aa, bb)
;	$LOW_NUM_STYLE_TRANSLITERATION(11), A transliteration module will be used to produce numbers in Chinese, Japanese, etc.
;	$LOW_NUM_STYLE_NATIVE_NUMBERING(12), The NativeNumberSupplier service will be called to produce numbers in native languages.
;	$LOW_NUM_STYLE_FULLWIDTH_ARABIC(13), Numbering for full width Arabic number.
;	$LOW_NUM_STYLE_CIRCLE_NUMBER(14), 	Bullet for Circle Number.
;	$LOW_NUM_STYLE_NUMBER_LOWER_ZH(15), Numbering for Chinese lower case number.
;	$LOW_NUM_STYLE_NUMBER_UPPER_ZH(16), Numbering for Chinese upper case number.
;	$LOW_NUM_STYLE_NUMBER_UPPER_ZH_TW(17), Numbering for Traditional Chinese upper case number.
;	$LOW_NUM_STYLE_TIAN_GAN_ZH(18), Bullet for Chinese Tian Gan.
;	$LOW_NUM_STYLE_DI_ZI_ZH(19), Bullet for Chinese Di Zi.
;	$LOW_NUM_STYLE_NUMBER_TRADITIONAL_JA(20), Numbering for Japanese traditional number.
;	$LOW_NUM_STYLE_AIU_FULLWIDTH_JA(21), Bullet for Japanese AIU fullwidth.
;	$LOW_NUM_STYLE_AIU_HALFWIDTH_JA(22), Bullet for Japanese AIU halfwidth.
;	$LOW_NUM_STYLE_IROHA_FULLWIDTH_JA(23), Bullet for Japanese IROHA fullwidth.
;	$LOW_NUM_STYLE_IROHA_HALFWIDTH_JA(24), Bullet for Japanese IROHA halfwidth.
;	$LOW_NUM_STYLE_NUMBER_UPPER_KO(25), Numbering for Korean upper case number.
;	$LOW_NUM_STYLE_NUMBER_HANGUL_KO(26), Numbering for Korean Hangul number.
;	$LOW_NUM_STYLE_HANGUL_JAMO_KO(27), Bullet for Korean Hangul Jamo.
;	$LOW_NUM_STYLE_HANGUL_SYLLABLE_KO(28), Bullet for Korean Hangul Syllable.
;	$LOW_NUM_STYLE_HANGUL_CIRCLED_JAMO_KO(29), Bullet for Korean Hangul Circled Jamo.
;	$LOW_NUM_STYLE_HANGUL_CIRCLED_SYLLABLE_KO(30), Bullet for Korean Hangul Circled Syllable.
;	$LOW_NUM_STYLE_CHARS_ARABIC(31), Numbering in Arabic alphabet letters.
;	$LOW_NUM_STYLE_CHARS_THAI(32), Numbering in Thai alphabet letters.
;	$LOW_NUM_STYLE_CHARS_HEBREW(33), Numbering in Hebrew alphabet letters.
;	$LOW_NUM_STYLE_CHARS_NEPALI(34), Numbering in Nepali alphabet letters.
;	$LOW_NUM_STYLE_CHARS_KHMER(35), Numbering in Khmer alphabet letters.
;	$LOW_NUM_STYLE_CHARS_LAO(36), Numbering in Lao alphabet letters.
;	$LOW_NUM_STYLE_CHARS_TIBETAN(37), Numbering in Tibetan/Dzongkha alphabet letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_BG(38), Numbering in Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_BG(39), Numbering in Cyrillic alphabet lower case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_BG(40), Numbering in Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_BG(41), Numbering in Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_RU(42), Numbering in Russian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_RU(43), Numbering in Russian Cyrillic alphabet lower case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_RU(44), Numbering in Russian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_RU(45), Numbering in Russian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_PERSIAN(46), Numbering in Persian alphabet letters.
;	$LOW_NUM_STYLE_CHARS_MYANMAR(47), Numbering in Myanmar alphabet letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_SR(48), Numbering in Serbian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_SR(49), Numbering in Russian Serbian alphabet lower case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_SR(50), Numbering in Serbian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_SR(51), Numbering in Serbian Cyrillic alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_GREEK_UPPER_LETTER(52), Numbering in Greek alphabet upper case letters.
;	$LOW_NUM_STYLE_CHARS_GREEK_LOWER_LETTER(53), Numbering in Greek alphabet lower case letters.
;	$LOW_NUM_STYLE_CHARS_ARABIC_ABJAD(54), Numbering in Arabic alphabet using abjad sequence.
;	$LOW_NUM_STYLE_CHARS_PERSIAN_WORD(55), Numbering in Persian words.
;	$LOW_NUM_STYLE_NUMBER_HEBREW(56), Numbering in Hebrew numerals.
;	$LOW_NUM_STYLE_NUMBER_ARABIC_INDIC(57), Numbering in Arabic-Indic numerals.
;	$LOW_NUM_STYLE_NUMBER_EAST_ARABIC_INDIC(58), Numbering in East Arabic-Indic numerals.
;	$LOW_NUM_STYLE_NUMBER_INDIC_DEVANAGARI(59), Numbering in Indic Devanagari numerals.
;	$LOW_NUM_STYLE_TEXT_NUMBER(60), Numbering in ordinal numbers of the language of the text node. (1st, 2nd, 3rd)
;	$LOW_NUM_STYLE_TEXT_CARDINAL(61), Numbering in cardinal numbers of the language of the text node. (One, Two)
;	$LOW_NUM_STYLE_TEXT_ORDINAL(62), Numbering in ordinal numbers of the language of the text node. (First, Second)
;	$LOW_NUM_STYLE_SYMBOL_CHICAGO(63), Footnoting symbols according the University of Chicago style.
;	$LOW_NUM_STYLE_ARABIC_ZERO(64), Numbering is in Arabic numbers, padded with zero to have a length of at least two. (01, 02)
;	$LOW_NUM_STYLE_ARABIC_ZERO3(65), Numbering is in Arabic numbers, padded with zero to have a length of at least three.
;	$LOW_NUM_STYLE_ARABIC_ZERO4(66), Numbering is in Arabic numbers, padded with zero to have a length of at least four.
;	$LOW_NUM_STYLE_ARABIC_ZERO5(67), Numbering is in Arabic numbers, padded with zero to have a length of at least five.
;	$LOW_NUM_STYLE_SZEKELY_ROVAS(68), Numbering is in Szekely rovas (Old Hungarian) numerals.
;	$LOW_NUM_STYLE_NUMBER_DIGITAL_KO(69), Numbering is in Korean Digital number.
;	$LOW_NUM_STYLE_NUMBER_DIGITAL2_KO(70), Numbering is in Korean Digital Number, reserved "koreanDigital2".
;	$LOW_NUM_STYLE_NUMBER_LEGAL_KO(71), Numbering is in Korean Legal Number, reserved "koreanLegal".
; Counting Type Constants: $LOW_FOOTNOTE_COUNT_PER_PAGE(0), Restarts the numbering of footnotes at the top of each page. This option is only available if End of Doc is set to False.
;							$LOW_FOOTNOTE_COUNT_PER_CHAP(1), Restarts the numbering of footnotes at the beginning of each chapter.
;							$LOW_FOOTNOTE_COUNT_PER_DOC(2), Numbers the footnotes in the document sequentially.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteSettingsAutoNumber(ByRef $oDoc, $iNumFormat = Null, $iStartAt = Null, $sBefore = Null, $sAfter = Null, $iCounting = Null, $bEndOfDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFNSettings[6]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iNumFormat, $iStartAt, $sBefore, $sAfter, $iCounting, $bEndOfDoc) Then
		__LOWriter_ArrayFill($avFNSettings, $oDoc.FootnoteSettings.NumberingType(), ($oDoc.FootnoteSettings.StartAt + 1), _
				$oDoc.FootnoteSettings.Prefix(), $oDoc.FootnoteSettings.Suffix(), $oDoc.FootnoteSettings.FootnoteCounting(), _
				$oDoc.FootnoteSettings.PositionEndOfDoc())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avFNSettings)
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDoc.FootnoteSettings.NumberingType = $iNumFormat
		$iError = ($oDoc.FootnoteSettings.NumberingType() = $iNumFormat) ? $iError : BitOR($iError, 1)
	EndIf

	; 0 Based -- Minus 1
	If ($iStartAt <> Null) Then
		If Not __LOWriter_IntIsBetween($iStartAt, 1, 9999) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDoc.FootnoteSettings.StartAt = ($iStartAt - 1)
		$iError = ($oDoc.FootnoteSettings.StartAt() = ($iStartAt - 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sBefore <> Null) Then
		If Not IsString($sBefore) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oDoc.FootnoteSettings.Prefix = $sBefore
		$iError = ($oDoc.FootnoteSettings.Prefix() = $sBefore) ? $iError : BitOR($iError, 4)
	EndIf

	If ($sAfter <> Null) Then
		If Not IsString($sAfter) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDoc.FootnoteSettings.Suffix = $sAfter
		$iError = ($oDoc.FootnoteSettings.Suffix() = $sAfter) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iCounting <> Null) Then
		If Not __LOWriter_IntIsBetween($iCounting, $LOW_FOOTNOTE_COUNT_PER_PAGE, $LOW_FOOTNOTE_COUNT_PER_DOC) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDoc.FootnoteSettings.FootnoteCounting = $iCounting
		$iError = ($oDoc.FootnoteSettings.FootnoteCounting() = $iCounting) ? $iError : BitOR($iError, 16)
	EndIf

	If ($bEndOfDoc <> Null) Then
		If Not IsBool($bEndOfDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oDoc.FootnoteSettings.PositionEndOfDoc = $bEndOfDoc
		$iError = ($oDoc.FootnoteSettings.PositionEndOfDoc() = $bEndOfDoc) ? $iError : BitOR($iError, 32)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnoteSettingsAutoNumber

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteSettingsContinuation
; Description ...: Set or Retrieve Footnote continuation settings.
; Syntax ........: _LOWriter_FootnoteSettingsContinuation(Byref $oDoc[, $sEnd = Null[, $sBegin = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sEnd                - [optional] a string value. Default is Null. The text to display at the end of a Footnote before it continues on the next page.
;                  $sBegin              - [optional] a string value. Default is Null. The text to display at the beginning of a Footnote that has continued on the next page.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sEnd not a String.
;				   @Error 1 @Extended 3 Return 0 = $sBegin not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sEnd
;				   |								2 = Error setting $sBegin
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteSettingsContinuation(ByRef $oDoc, $sEnd = Null, $sBegin = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asFNSettings[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sEnd, $sBegin) Then
		__LOWriter_ArrayFill($asFNSettings, $oDoc.FootnoteSettings.EndNotice(), $oDoc.FootnoteSettings.BeginNotice())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $asFNSettings)
	EndIf

	If ($sEnd <> Null) Then
		If Not IsString($sEnd) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDoc.FootnoteSettings.EndNotice = $sEnd
		$iError = ($oDoc.FootnoteSettings.EndNotice() = $sEnd) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sBegin <> Null) Then
		If Not IsString($sBegin) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDoc.FootnoteSettings.BeginNotice = $sBegin
		$iError = ($oDoc.FootnoteSettings.BeginNotice() = $sBegin) ? $iError : BitOR($iError, 2)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnoteSettingsContinuation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteSettingsStyles
; Description ...: Set or Retrieve Document Footnote Style settings.
; Syntax ........: _LOWriter_FootnoteSettingsStyles(Byref $oDoc[, $sParagraph = Null[, $sPage = Null[, $sTextArea = Null[, $sFootnoteArea = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sParagraph          - [optional] a string value. Default is Null. The Footnote Text Paragraph Style.
;                  $sPage               - [optional] a string value. Default is Null. The Page Style to use for the Footnote pages.
;				   +						Only valid if the footnotes are set to End of Document, instead of per page.
;                  $sTextArea           - [optional] a string value. Default is Null. The Character Style to use for the Footnote anchor in the document text.
;                  $sFootnoteArea       - [optional] a string value. Default is Null. The Character Style to use for the Footnote number in the footnote text.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sParagraph not a String.
;				   @Error 1 @Extended 3 Return 0 = Paragraph Style referenced in $sParagraph not found in Document.
;				   @Error 1 @Extended 4 Return 0 = $sPage not a String.
;				   @Error 1 @Extended 5 Return 0 = Page Style referenced in $sPage not found in Document.
;				   @Error 1 @Extended 6 Return 0 = $sTextArea not a String.
;				   @Error 1 @Extended 7 Return 0 = Character Style referenced in $sTextArea not found in Document.
;				   @Error 1 @Extended 8 Return 0 = $sFootnoteArea not a String.
;				   @Error 1 @Extended 9 Return 0 = Character Style referenced in $sFootnoteArea not found in Document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sParagraph
;				   |								2 = Error setting $sPage
;				   |								4 = Error setting $sTextArea
;				   |								8 = Error setting $sFootnoteArea
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStylesGetNames, _LOWriter_PageStylesGetNames, _LOWriter_CharStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteSettingsStyles(ByRef $oDoc, $sParagraph = Null, $sPage = Null, $sTextArea = Null, $sFootnoteArea = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFNSettings[4]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sParagraph, $sPage, $sTextArea, $sFootnoteArea) Then
		__LOWriter_ArrayFill($avFNSettings, __LOWriter_ParStyleNameToggle($oDoc.FootnoteSettings.ParaStyleName(), True), _
				__LOWriter_PageStyleNameToggle($oDoc.FootnoteSettings.PageStyleName(), True), _
				__LOWriter_CharStyleNameToggle($oDoc.FootnoteSettings.AnchorCharStyleName(), True), _
				__LOWriter_CharStyleNameToggle($oDoc.FootnoteSettings.CharStyleName(), True))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avFNSettings)
	EndIf

	If ($sParagraph <> Null) Then
		If Not IsString($sParagraph) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		If Not _LOWriter_ParStyleExists($oDoc, $sParagraph) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$sParagraph = __LOWriter_ParStyleNameToggle($sParagraph)
		$oDoc.FootnoteSettings.ParaStyleName = $sParagraph
		$iError = ($oDoc.FootnoteSettings.ParaStyleName() = $sParagraph) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sPage <> Null) Then
		If Not IsString($sPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		If Not _LOWriter_PageStyleExists($oDoc, $sPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$sPage = __LOWriter_PageStyleNameToggle($sPage)
		$oDoc.FootnoteSettings.PageStyleName = $sPage
		$iError = ($oDoc.FootnoteSettings.PageStyleName() = $sPage) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sTextArea <> Null) Then
		If Not IsString($sTextArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_CharStyleExists($oDoc, $sTextArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$sTextArea = __LOWriter_CharStyleNameToggle($sTextArea)
		$oDoc.FootnoteSettings.AnchorCharStyleName = $sTextArea
		$iError = ($oDoc.FootnoteSettings.AnchorCharStyleName() = $sTextArea) ? $iError : BitOR($iError, 4)
	EndIf

	If ($sFootnoteArea <> Null) Then
		If Not IsString($sFootnoteArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		If Not _LOWriter_CharStyleExists($oDoc, $sFootnoteArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$sFootnoteArea = __LOWriter_CharStyleNameToggle($sFootnoteArea)
		$oDoc.FootnoteSettings.CharStyleName = $sFootnoteArea
		$iError = ($oDoc.FootnoteSettings.CharStyleName() = $sFootnoteArea) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnoteSettingsStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnotesGetList
; Description ...: Retrieve an array of Footnote Objects contained in a Document.
; Syntax ........: _LOWriter_FootnotesGetList(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Footnotes Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully searched for Footnotes, none contained in document.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for Footnotes, Returning Array of Footnote Objects. @Extended set to number found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FootnoteDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnotesGetList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFootNotes
	Local $aoFootnotes[0]
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oFootNotes = $oDoc.getFootnotes()
	If Not IsObj($oFootNotes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$iCount = $oFootNotes.getCount()

	If ($iCount > 0) Then
		ReDim $aoFootnotes[$iCount]

		For $i = 0 To $iCount - 1
			$aoFootnotes[$i] = $oFootNotes.getByIndex($i)
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return ($iCount > 0) ? SetError($__LOW_STATUS_SUCCESS, $iCount, $aoFootnotes) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnotesGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_SearchDescriptorCreate
; Description ...: Create a Search Descriptor for searching a document.
; Syntax ........: _LOWriter_SearchDescriptorCreate(Byref $oDoc[, $bBackwards = False[, $bMatchCase = False[, $bWholeWord = False[, $bRegExp = False[, $bStyles = False[, $bSearchPropValues = False]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bBackwards          - [optional] a boolean value. Default is False. If True, the document is searched backwards.
;                  $bMatchCase          - [optional] a boolean value. Default is False. If True, the case of the letters is important for the Search.
;                  $bWholeWord          - [optional] a boolean value. Default is False. If True, only complete words will be found.
;                  $bRegExp             - [optional] a boolean value. Default is False. If True, the search string is evaluated as a regular expression.
;                  $bStyles             - [optional] a boolean value. Default is False. If True, the string is considered a Paragraph Style name,
;				   +						and the search will return any paragraph utilizing the specified name, EXCEPT if you input Format properties to search for,
;				   +						then setting this to True causes the search to search both for direct formatting matching those properties
;				   +						and also Paragraph/Character styles that contain matching properties.
;                  $bSearchPropValues   - [optional] a boolean value. Default is False. If True, any formatting properties searched for are matched based on their value,
;				   +						else if false, the search only looks for their existence. See Remarks.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bBackwards not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bMatchCase not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bWholeWord not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bRegExp not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $bStyles not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $bSearchPropValues not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create Search Descriptor.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returns a Search Descriptor Object for setting Search options.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bSearchPropValues is equivalent to the difference in selecting "Format" options in Libre Office's search box and "Attributes".
;						Setting $bSearchPropValues to True, means that the search will look for matches using the specified property AND having the specified value,
;						such as Character Weight, Bold, only matches that have Character weight of Bold will be returned, whereas if $bSearchPropValues is set to false,
;						the search only looks for matches that have the specified property, regardless of its value.
;						Such as Character weight, would match Bold, Semi-Bold, etc. From my understanding, the search is based on anything directly formatted unless $bStyles is also true.
;					Note: The returned Search Descriptor is only good for the Document it was created by, it WILL NOT work for
;						other documents.
; Related .......: _LOWriter_SearchDescriptorModify, _LOWriter_SearchDescriptorSimilarityModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_SearchDescriptorCreate(ByRef $oDoc, $bBackwards = False, $bMatchCase = False, $bWholeWord = False, $bRegExp = False, $bStyles = False, $bSearchPropValues = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSrchDescript

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If Not IsBool($bBackwards) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bMatchCase) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bWholeWord) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bRegExp) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bStyles) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bSearchPropValues) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)

	$oSrchDescript = $oDoc.createSearchDescriptor()
	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	With $oSrchDescript
		.SearchBackwards = $bBackwards
		.SearchCaseSensitive = $bMatchCase
		.SearchWords = $bWholeWord
		.SearchRegularExpression = $bRegExp
		.SearchStyles = $bStyles
		.ValueSearch = $bSearchPropValues
	EndWith

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oSrchDescript)
EndFunc   ;==>_LOWriter_SearchDescriptorCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_SearchDescriptorModify
; Description ...: Modify Search Descriptor settings of an existing Search Descriptor Object.
; Syntax ........: _LOWriter_SearchDescriptorModify(Byref $oSrchDescript[, $bBackwards = Null[, $bMatchCase = Null[, $bWholeWord = Null[, $bRegExp = Null[, $bStyles = Null[, $bSearchPropValues = Null]]]]]])
; Parameters ....: $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $bBackwards          - [optional] a boolean value. Default is False. If True, the document is searched backwards.
;                  $bMatchCase          - [optional] a boolean value. Default is False. If True, the case of the letters is important for the Search.
;                  $bWholeWord          - [optional] a boolean value. Default is False. If True, only complete words will be found.
;                  $bRegExp             - [optional] a boolean value. Default is False. If True, the search string is evaluated as a regular expression. Cannot be set to True if Similarity Search is set to True.
;                  $bStyles             - [optional] a boolean value. Default is False. If True, the string is considered a Paragraph Style name,
;				   +						and the search will return any paragraph utilizing the specified name, EXCEPT if you input Format properties to search for,
;				   +						then setting this to True causes the search to search both for direct formatting matching those properties
;				   +						and also Paragraph/Character styles that contain matching properties.
;                  $bSearchPropValues   - [optional] a boolean value. Default is False. If True, any formatting properties searched for are matched based on their value, else if false,
;				   +						the search only looks for their existence. See Remarks.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript Object not a Search Descriptor Object.
;				   @Error 1 @Extended 3 Return 0 = $bBackwards not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bMatchCase not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bWholeWord not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $bRegExp not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $bStyles not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $bSearchPropValues not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = $bRegExp is set to True while Similarity Search is also set to True.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Returns 1 after directly modifying Search Descriptor Object.
; ;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bSearchPropValues is equivalent to the difference in selecting "Format" options in Libre Office's search box and "Attributes".
;						Setting $bSearchPropValues to True, means that the search will look for matches using the specified property AND having the specified value,
;						such as Character Weight, Bold, only matches that have Character weight of Bold will be returned, whereas if $bSearchPropValues is set to false,
;						the search only looks for matches that have the specified property, regardless of its value.
;						Such as Character weight, would match Bold, Semi-Bold, etc. From my understanding, the search is based on anything directly formatted unless $bStyles is also true.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;						Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_SearchDescriptorCreate, _LOWriter_SearchDescriptorSimilarityModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_SearchDescriptorModify(ByRef $oSrchDescript, $bBackwards = Null, $bMatchCase = Null, $bWholeWord = Null, $bRegExp = Null, $bStyles = Null, $bSearchPropValues = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSrchDescript[6]

	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bBackwards, $bMatchCase, $bWholeWord, $bRegExp, $bStyles, $bSearchPropValues) Then
		__LOWriter_ArrayFill($avSrchDescript, $oSrchDescript.SearchBackwards(), $oSrchDescript.SearchCaseSensitive(), $oSrchDescript.SearchWords(), _
				$oSrchDescript.SearchRegularExpression(), $oSrchDescript.SearchStyles(), $oSrchDescript.getValueSearch())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avSrchDescript)
	EndIf

	If ($bBackwards <> Null) Then
		If Not IsBool($bBackwards) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oSrchDescript.SearchBackwards = $bBackwards
	EndIf

	If ($bMatchCase <> Null) Then
		If Not IsBool($bMatchCase) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oSrchDescript.SearchCaseSensitive = $bMatchCase
	EndIf

	If ($bWholeWord <> Null) Then
		If Not IsBool($bWholeWord) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oSrchDescript.SearchWords = $bWholeWord
	EndIf

	If ($bRegExp <> Null) Then
		If Not IsBool($bRegExp) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If ($bRegExp = True) And ($oSrchDescript.SearchSimilarity = True) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		$oSrchDescript.SearchRegularExpression = $bRegExp
	EndIf

	If ($bStyles <> Null) Then
		If Not IsBool($bStyles) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oSrchDescript.SearchStyles = $bStyles
	EndIf

	If ($bSearchPropValues <> Null) Then
		If Not IsBool($bSearchPropValues) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oSrchDescript.ValueSearch = $bSearchPropValues
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_SearchDescriptorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_SearchDescriptorSimilarityModify
; Description ...: Modify Similarity Search Settings for an existing Search Descriptor Object.
; Syntax ........: _LOWriter_SearchDescriptorSimilarityModify(Byref $oSrchDescript[, $bSimilarity = Null[, $bCombine = Null[, $iRemove = Null[, $iAdd = Null[, $iExchange = Null]]]]])
; Parameters ....: $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $bSimilarity         - [optional] a boolean value. Default is Null. If True, a "similarity search" is performed.
;                  $bCombine            - [optional] a boolean value. Default is Null. If True, all similarity rules ($iRemove, $iAdd, and $iExchange) are applied together.
;                  $iRemove             - [optional] an integer value. Default is Null. Specifies the number of characters that may be ignored to match the search pattern.
;                  $iAdd                - [optional] an integer value. Default is Null. Specifies the number of characters that must be added to match the search pattern.
;                  $iExchange           - [optional] an integer value. Default is Null. Specifies the number of characters that must be replaced to match the search pattern.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript Object not a Search Descriptor Object.
;				   @Error 1 @Extended 3 Return 0 = $bSimilarity not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bCombine not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iRemove, $iAdd, or $iExchange set to a value, but $bSimilarity not set to True.
;				   @Error 1 @Extended 6 Return 0 = $iRemove not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iAdd not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iExchange not an Integer.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = $bSimilarity is set to True while Regular Expression Search is also set to True.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Returns 1 after directly modifying Search Descriptor Object.
; ;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_SearchDescriptorCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_SearchDescriptorSimilarityModify(ByRef $oSrchDescript, $bSimilarity = Null, $bCombine = Null, $iRemove = Null, $iAdd = Null, $iExchange = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSrchDescript[5]

	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bSimilarity, $bCombine, $iRemove, $iAdd, $iExchange) Then
		__LOWriter_ArrayFill($avSrchDescript, $oSrchDescript.SearchSimilarity(), $oSrchDescript.SearchSimilarityRelax(), _
				$oSrchDescript.SearchSimilarityRemove(), $oSrchDescript.SearchSimilarityAdd(), $oSrchDescript.SearchSimilarityExchange())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avSrchDescript)
	EndIf

	If ($bSimilarity <> Null) Then
		If Not IsBool($bSimilarity) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		If ($bSimilarity = True) And ($oSrchDescript.SearchRegularExpression = True) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		$oSrchDescript.SearchSimilarity = $bSimilarity
	EndIf

	If ($bCombine <> Null) Then
		If Not IsBool($bCombine) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oSrchDescript.SearchSimilarityRelax = $bCombine
	EndIf

	If Not __LOWriter_VarsAreNull($iRemove, $iAdd, $iExchange) Then
		If ($oSrchDescript.SearchSimilarity() = False) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		If ($iRemove <> Null) Then
			If Not IsInt($iRemove) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			$oSrchDescript.SearchSimilarityRemove = $iRemove
		EndIf

		If ($iAdd <> Null) Then
			If Not IsInt($iAdd) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
			$oSrchDescript.SearchSimilarityAdd = $iAdd
		EndIf

		If ($iExchange <> Null) Then
			If Not IsInt($iExchange) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
			$oSrchDescript.SearchSimilarityExchange = $iExchange
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_SearchDescriptorSimilarityModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapesGetNames
; Description ...: Return a list of Shape names contained in a document.
; Syntax ........: _LOWriter_ShapesGetNames(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object.  A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 2D Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Shapes Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning 2D Array containing a list of Shape names contained contained in a document,
;				   +									the first column ($aArray[0][0] contains the shape name,
;				   +									the second column ($aArray[0][1] contains the shape's Implementation name. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Implementation name identifies what type of shape object it is, as there can be multiple things counted as "Shapes", such as Text Frames etc.
;					I have found the three Implementation names being returned, SwXTextFrame, indicating the shape is actually a Text Frame, SwXShape, is a regular shape such as a line, circle etc.
;					And "SwXTextGraphicObject", which is an image / picture. There may be other return types I haven't found yet.
;					Images inserted into the document are also listed as TextFrames in the shapes category.
;					There isn't and easy way to differentiate between them yet, see _LOWriter_FramesGetNames, to search for Frames in the shapes category.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapesGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asShapeNames[0][2]
	Local $oShapes

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oShapes = $oDoc.DrawPage()
	If Not IsObj($oShapes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $oShapes.hasElements() Then
		ReDim $asShapeNames[$oShapes.getCount()][2]
		For $i = 0 To $oShapes.getCount() - 1
			$asShapeNames[$i][0] = $oShapes.getByIndex($i).Name()
			If $oShapes.getByIndex($i).supportsService("com.sun.star.drawing.Text") Then
				; If Supports Text Method, then get that impl. name, else just te regular impl. name.
				$asShapeNames[$i][1] = $oShapes.getByIndex($i).Text.ImplementationName()
			Else
				$asShapeNames[$i][1] = $oShapes.getByIndex($i).ImplementationName()
			EndIf

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, UBound($asShapeNames), $asShapeNames)
EndFunc   ;==>_LOWriter_ShapesGetNames
