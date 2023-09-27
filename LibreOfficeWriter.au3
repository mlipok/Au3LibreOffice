;~ #AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

#include "LibreOfficeWriter_Doc.au3"
#include "LibreOfficeWriter_Frame.au3"
#include "LibreOfficeWriter_Table.au3"

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter)
; AutoIt Version : v3.3.16.1
; UDF Version    : 0.0.0.3
; Description ...: Provides basic functionality through Autoit for interacting with Libre Office Writer.
; Author(s) .....: donnyh13
; Sources . . . .:  jguinch -- Printmgr.au3, used (_PrintMgr_EnumPrinter);
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
;_LOWriter_CellBackColor
;_LOWriter_CellBorderColor
;_LOWriter_CellBorderPadding
;_LOWriter_CellBorderStyle
;_LOWriter_CellBorderWidth
;_LOWriter_CellCreateTextCursor
;_LOWriter_CellFormula
;_LOWriter_CellGetDataType
;_LOWriter_CellGetError
;_LOWriter_CellGetName
;_LOWriter_CellProtect
;_LOWriter_CellString
;_LOWriter_CellValue
;_LOWriter_CellVertOrient
;_LOWriter_DateFormatKeyCreate
;_LOWriter_DateFormatKeyDelete
;_LOWriter_DateFormatKeyExists
;_LOWriter_DateFormatKeyGetString
;_LOWriter_DateFormatKeyList
;_LOWriter_DateStructCreate
;_LOWriter_DateStructModify
;_LOWriter_DirFrmtCharBorderColor
;_LOWriter_DirFrmtCharBorderPadding
;_LOWriter_DirFrmtCharBorderStyle
;_LOWriter_DirFrmtCharBorderWidth
;_LOWriter_DirFrmtCharEffect
;_LOWriter_DirFrmtCharPosition
;_LOWriter_DirFrmtCharRotateScale
;_LOWriter_DirFrmtCharShadow
;_LOWriter_DirFrmtCharSpacing
;_LOWriter_DirFrmtClear
;_LOWriter_DirFrmtFont
;_LOWriter_DirFrmtFontColor
;_LOWriter_DirFrmtGetCurStyles
;_LOWriter_DirFrmtOverLine
;_LOWriter_DirFrmtParAlignment
;_LOWriter_DirFrmtParBackColor
;_LOWriter_DirFrmtParBorderColor
;_LOWriter_DirFrmtParBorderPadding
;_LOWriter_DirFrmtParBorderStyle
;_LOWriter_DirFrmtParBorderWidth
;_LOWriter_DirFrmtParDropCaps
;_LOWriter_DirFrmtParHyphenation
;_LOWriter_DirFrmtParIndent
;_LOWriter_DirFrmtParOutLineAndList
;_LOWriter_DirFrmtParPageBreak
;_LOWriter_DirFrmtParShadow
;_LOWriter_DirFrmtParSpace
;_LOWriter_DirFrmtParTabStopCreate
;_LOWriter_DirFrmtParTabStopDelete
;_LOWriter_DirFrmtParTabStopList
;_LOWriter_DirFrmtParTabStopMod
;_LOWriter_DirFrmtParTxtFlowOpt
;_LOWriter_DirFrmtStrikeOut
;_LOWriter_DirFrmtUnderLine
;_LOWriter_EndnoteDelete
;_LOWriter_EndnoteGetAnchor
;_LOWriter_EndnoteGetTextCursor
;_LOWriter_EndnoteInsert
;_LOWriter_EndnoteModifyAnchor
;_LOWriter_EndnoteSettingsAutoNumber
;_LOWriter_EndnoteSettingsStyles
;_LOWriter_EndnotesGetList
;_LOWriter_FieldAuthorInsert
;_LOWriter_FieldAuthorModify
;_LOWriter_FieldChapterInsert
;_LOWriter_FieldChapterModify
;_LOWriter_FieldCombCharInsert
;_LOWriter_FieldCombCharModify
;_LOWriter_FieldCommentInsert
;_LOWriter_FieldCommentModify
;_LOWriter_FieldCondTextInsert
;_LOWriter_FieldCondTextModify
;_LOWriter_FieldCurrentDisplayGet
;_LOWriter_FieldDateTimeInsert
;_LOWriter_FieldDateTimeModify
;_LOWriter_FieldDelete
;_LOWriter_FieldDocInfoCommentsInsert
;_LOWriter_FieldDocInfoCommentsModify
;_LOWriter_FieldDocInfoCreateAuthInsert
;_LOWriter_FieldDocInfoCreateAuthModify
;_LOWriter_FieldDocInfoCreateDateTimeInsert
;_LOWriter_FieldDocInfoCreateDateTimeModify
;_LOWriter_FieldDocInfoEditTimeInsert
;_LOWriter_FieldDocInfoEditTimeModify
;_LOWriter_FieldDocInfoKeywordsInsert
;_LOWriter_FieldDocInfoKeywordsModify
;_LOWriter_FieldDocInfoModAuthInsert
;_LOWriter_FieldDocInfoModAuthModify
;_LOWriter_FieldDocInfoModDateTimeInsert
;_LOWriter_FieldDocInfoModDateTimeModify
;_LOWriter_FieldDocInfoPrintAuthInsert
;_LOWriter_FieldDocInfoPrintAuthModify
;_LOWriter_FieldDocInfoPrintDateTimeInsert
;_LOWriter_FieldDocInfoPrintDateTimeModify
;_LOWriter_FieldDocInfoRevNumInsert
;_LOWriter_FieldDocInfoRevNumModify
;_LOWriter_FieldDocInfoSubjectInsert
;_LOWriter_FieldDocInfoSubjectModify
;_LOWriter_FieldDocInfoTitleInsert
;_LOWriter_FieldDocInfoTitleModify
;_LOWriter_FieldFileNameInsert
;_LOWriter_FieldFileNameModify
;_LOWriter_FieldFuncHiddenParInsert
;_LOWriter_FieldFuncHiddenParModify
;_LOWriter_FieldFuncHiddenTextInsert
;_LOWriter_FieldFuncHiddenTextModify
;_LOWriter_FieldFuncInputInsert
;_LOWriter_FieldFuncInputModify
;_LOWriter_FieldFuncPlaceholderInsert
;_LOWriter_FieldFuncPlaceholderModify
;_LOWriter_FieldGetAnchor
;_LOWriter_FieldInputListInsert
;_LOWriter_FieldInputListModify
;_LOWriter_FieldPageNumberInsert
;_LOWriter_FieldPageNumberModify
;_LOWriter_FieldRefBookMarkInsert
;_LOWriter_FieldRefBookMarkModify
;_LOWriter_FieldRefEndnoteInsert
;_LOWriter_FieldRefEndnoteModify
;_LOWriter_FieldRefFootnoteInsert
;_LOWriter_FieldRefFootnoteModify
;_LOWriter_FieldRefGetType
;_LOWriter_FieldRefInsert
;_LOWriter_FieldRefMarkDelete
;_LOWriter_FieldRefMarkGetAnchor
;_LOWriter_FieldRefMarkList
;_LOWriter_FieldRefMarkSet
;_LOWriter_FieldRefModify
;_LOWriter_FieldsAdvGetList
;_LOWriter_FieldsDocInfoGetList
;_LOWriter_FieldSenderInsert
;_LOWriter_FieldSenderModify
;_LOWriter_FieldSetVarInsert
;_LOWriter_FieldSetVarMasterCreate
;_LOWriter_FieldSetVarMasterDelete
;_LOWriter_FieldSetVarMasterExists
;_LOWriter_FieldSetVarMasterGetObj
;_LOWriter_FieldSetVarMasterList
;_LOWriter_FieldSetVarMasterListFields
;_LOWriter_FieldSetVarModify
;_LOWriter_FieldsGetList
;_LOWriter_FieldShowVarInsert
;_LOWriter_FieldShowVarModify
;_LOWriter_FieldStatCountInsert
;_LOWriter_FieldStatCountModify
;_LOWriter_FieldStatTemplateInsert
;_LOWriter_FieldStatTemplateModify
;_LOWriter_FieldUpdate
;_LOWriter_FieldVarSetPageInsert
;_LOWriter_FieldVarSetPageModify
;_LOWriter_FieldVarShowPageInsert
;_LOWriter_FieldVarShowPageModify
;_LOWriter_FindFormatModifyAlignment
;_LOWriter_FindFormatModifyEffects
;_LOWriter_FindFormatModifyFont
;_LOWriter_FindFormatModifyHyphenation
;_LOWriter_FindFormatModifyIndent
;_LOWriter_FindFormatModifyOverline
;_LOWriter_FindFormatModifyPageBreak
;_LOWriter_FindFormatModifyPosition
;_LOWriter_FindFormatModifyRotateScaleSpace
;_LOWriter_FindFormatModifySpacing
;_LOWriter_FindFormatModifyStrikeout
;_LOWriter_FindFormatModifyTxtFlowOpt
;_LOWriter_FindFormatModifyUnderline
;_LOWriter_FootnoteDelete
;_LOWriter_FootnoteGetAnchor
;_LOWriter_FootnoteGetTextCursor
;_LOWriter_FootnoteInsert
;_LOWriter_FootnoteModifyAnchor
;_LOWriter_FootnoteSettingsAutoNumber
;_LOWriter_FootnoteSettingsContinuation
;_LOWriter_FootnoteSettingsStyles
;_LOWriter_FootnotesGetList
;_LOWriter_FormatKeyCreate
;_LOWriter_FormatKeyDelete
;_LOWriter_FormatKeyExists
;_LOWriter_FormatKeyGetString
;_LOWriter_FormatKeyList
;_LOWriter_SearchDescriptorCreate
;_LOWriter_SearchDescriptorModify
;_LOWriter_SearchDescriptorSimilarityModify
;_LOWriter_ShapesGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBackColor
; Description ...: Set and Retrieve the Background color of a Cell or Cell Range.
; Syntax ........: _LOWriter_CellBackColor(Byref $oCell[, $iBackColor = Null[, $bBackTransparent = Null]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from any Table Cell
;				   +						Object creation or retrieval functions.
;                  $iBackColor          - [optional] an integer value. Default is Null. Specify the Cell background color as
;				   +						a Long Integer. See Remarks. Set to $LOW_COLOR_OFF(-1) to disable Background color.
;                  $bBackTransparent    - [optional] a boolean value. Default is Null. If True, the background color is
;				   +						transparent.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell variable not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iBackColor not an Integer, set less than -1 or greater than 16777215.
;				   @Error 1 @Extended 3 Return 0 = $bBackTransparent not a Boolean and not set to Null.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $iBackColor
;				   |								2 = Error setting $bBackTransparent
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					$iBackColor is set using Long integer. See _LOWriter_ConvertColorToLong,
;						_LOWriter_ConvertColorFromLong. There are also preset colors, listed below.
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
		$iError = ($oCell.BackColor() = $iBackColor) ? $iError : BitOR($iError, 1) ;Error setting color.
	EndIf

	If ($bBackTransparent <> Null) Then
		If Not IsBool($bBackTransparent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oCell.BackTransparent = $bBackTransparent
		$iError = ($oCell.BackTransparent() = $bBackTransparent) ? $iError : BitOR($iError, 2) ;Error setting BackTransparent.
	EndIf

	Return ($iError = 0) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0)
EndFunc   ;==>_LOWriter_CellBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellBorderColor
; Description ...: Set the Cell or Cell Range Border Line Color. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_CellBorderColor(Byref $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from any Table Cell
;				   +						Object creation or retrieval functions.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Color of the
;				   +						Cell in Long Color code format. One of the predefined constants listed below can be
;				   +						used, or a custom value.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Color of the
;				   +						Cell in Long Color code format. One of the predefined constants listed below can be
;				   +						used, or a custom value.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Color of the
;				   +						Cell in Long Color code format. One of the predefined constants listed below can be
;				   +						used, or a custom value.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Color of the
;				   +						Cell in Long Color code format. One of the predefined constants listed below can be
;				   +						used, or a custom value.
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
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
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
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from any Table Cell
;				   +						Object creation or retrieval functions.
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Distance between the Border
;				   +						and Cell text in Micrometers(uM).
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Distance between the
;				   +						Border and Cell text in Micrometers(uM).
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Distance between the Border
;				   +						and Cell text in Micrometers(uM).
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Distance between the Border
;				   +						and Cell text in Micrometers(uM).
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell parameter not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iTop not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iBottom not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $Left not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iRight not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $iTop border distance
;				   |								2 = Error setting $iBottom border distance
;				   |								4 = Error setting $iLeft border distance
;				   |								8 = Error setting $iRight border distance
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
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
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from any Table Cell
;				   +						Object creation or retrieval functions.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Style of the
;				   +							Cell using one of the line style constants, See below for list.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Style of the
;				   +							Cell using one of the line style constants, See below for list.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Style of the
;				   +							Cell using one of the line style constants, See below for list.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Style of the
;				   +							Cell using one of the line style constants, See below for list.
; Internal Remark: Error values for Initialization and Processing are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell Variable not Object type variable.
;				   @Error 1 @Extended 2 Return 0 = $iTop not an integer, or set to higher than 17 and not equal to 0x7FFF,
;				   +									Or $iTop is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 3 Return 0 = $iBottom not an integer, or set to higher than 17 and not equal to
;				   +								0x7FFF, Or $iBottom is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iLeft not an integer, or set to higher than 17 and not equal to 0x7FFF,
;				   +									Or $iLeft is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iRight not an integer, or set to higher than 17 and not equal to
;				   +									0x7FFF, Or $iRight is set to less than 0 or not set to Null.
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
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Style Constants: $LOW_BORDERSTYLE_NONE(0x7FFF) No border line,
;					$LOW_BORDERSTYLE_SOLID(0) Solid border line,
;					$LOW_BORDERSTYLE_DOTTED(1) Dotted border line,
;					$LOW_BORDERSTYLE_DASHED(2) Dashed border line,
;					$LOW_BORDERSTYLE_DOUBLE(3) Double border line,
;					$LOW_BORDERSTYLE_THINTHICK_SMALLGAP(4) Double border line with a thin line outside and a thick line inside
;						separated by a small gap,
;					$LOW_BORDERSTYLE_THINTHICK_MEDIUMGAP(5) Double border line with a thin line outside and a thick line inside
;						separated by a medium gap,
;					$LOW_BORDERSTYLE_THINTHICK_LARGEGAP(6) Double border line with a thin line outside and a thick line inside
;						separated by a large gap,
;					$LOW_BORDERSTYLE_THICKTHIN_SMALLGAP(7) Double border line with a thick line outside and a thin line inside
;						separated by a small gap,
;					$LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP(8) Double border line with a thick line outside and a thin line inside
;						separated by a medium gap,
;					$LOW_BORDERSTYLE_THICKTHIN_LARGEGAP(9) Double border line with a thick line outside and a thin line inside
;						separated by a large gap,
;					$LOW_BORDERSTYLE_EMBOSSED(10) 3D embossed border line,
;					$LOW_BORDERSTYLE_ENGRAVED(11) 3D engraved border line,
;					$LOW_BORDERSTYLE_OUTSET(12) Outset border line,
;					$LOW_BORDERSTYLE_INSET(13) Inset border line,
;					$LOW_BORDERSTYLE_FINE_DASHED(14) Finely dashed border line,
;					$LOW_BORDERSTYLE_DOUBLE_THIN(15) Double border line consisting of two fixed thin lines separated by a
;						variable gap,
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
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set them to 0
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Width Constants:  $LOW_BORDERWIDTH_HAIRLINE(2),
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
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object
;				   +						creation or retrieval functions.
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
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
;					_LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellCreateTextCursor(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ;Can only create a Text Cursor for individual cells.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.Text.createTextCursor())
EndFunc   ;==>_LOWriter_CellCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellFormula
; Description ...: Set or retrieve a formula for a cell.
; Syntax ........: _LOWriter_CellFormula(Byref $oCell[, $sFormula = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object
;				   +						creation or retrieval functions.
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
; Remarks .......: Note: Formula can only be set for an individual cell, not a range.
;					Setting the formula will overwrite any existing data in the cell.
; 					Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
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
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ;Can only set/get formula value for individual cells.
	If ($sFormula = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oCell.getFormula())

	$oCell.setFormula($sFormula)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CellFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellGetDataType
; Description ...: Get the Data type of a specific cell, see remarks.
; Syntax ........: _LOWriter_CellGetDataType(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object
;				   +						creation or retrieval functions.
; Return values .: Success: A Number.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a CellRange not an individual cell.
;				   --Success--
;				   @Error 0 @Extended 0 Return Number = Success. The Data Type in Number format, see constants below.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Returns the data type as one of the below constants, Note: If the data was entered by the keyboard, it is
;					generally recognized as a string regardless of the data contained.
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
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ;Can only get Data Type for individual cells

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.getType())
EndFunc   ;==>_LOWriter_CellGetDataType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellGetError
; Description ...: Get the formula error Value.
; Syntax ........: _LOWriter_CellGetError(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object
;				   +						creation or retrieval functions.
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
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ;Can only get Error for individual cells.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.getError())

EndFunc   ;==>_LOWriter_CellGetError

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellGetName
; Description ...: Retrieve the current Cell's name.
; Syntax ........: _LOWriter_CellGetName(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object
;				   +						creation or retrieval functions.
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
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ;Can only get Cell Name for individual cells.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.CellName())
EndFunc   ;==>_LOWriter_CellGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellProtect
; Description ...: Write-Protect a Cell
; Syntax ........: _LOWriter_CellProtect(Byref $oCell[, $bProtect = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object
;				   +						creation or retrieval functions.
;                  $bProtect            - [optional] a boolean value. Default is Null. True = Protected from Writing, False =
;				   +						Unprotected. See remarks.
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
;				   @Error 0 @Extended 0 Return Boolean = Success. $bProtect is set to Null, return will be the current setting
;				   +										of write-protection for the cell, a Boolean value.
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
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ;Can only set individual cell protect property.
	If ($bProtect = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.IsProtected())
	If Not IsBool($bProtect) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oCell.IsProtected = $bProtect
	Return ($oCell.IsProtected() = $bProtect) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)

EndFunc   ;==>_LOWriter_CellProtect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellString
; Description ...: Set or retrieve the current string for a cell.
; Syntax ........: _LOWriter_CellString(Byref $oCell[, $sString = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object
;				   +						creation or retrieval functions.
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
; Remarks .......: Note: String can only be set for an individual cell, not a range.
;					Setting the String will overwrite any existing data in the cell.
; 					Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellString(ByRef $oCell, $sString = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sString) And Not ($sString = Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ;Can only set/get a String for individual cells.

	If ($sString = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oCell.getString())

	$oCell.setString($sString)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CellString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CellValue
; Description ...: Set or retrieve a Numerical value to a Cell
; Syntax ........: _LOWriter_CellValue(Byref $oCell[, $nValue = Null])
; Parameters ....: $oCell               - [in/out] an object. A Table Cell Object returned from any Table Cell Object
;				   +						creation or retrieval functions.
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
; Remarks .......: Note: Value can only be set for an individual cell, not a range.
;					Setting the Value will overwrite any existing data in the cell.
;					For a value cell the value is returned, for a string cell zero is returned and for a formula cell the result
;						value of a formula is returned.
; 					Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
; Related .......:_LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, _LOWriter_TableGetCellObjByPosition,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CellValue(ByRef $oCell, $nValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsNumber($nValue) And Not ($nValue = Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If __LOWriter_IsCellRange($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ;Can only set/get individual cell values.

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
;				   @Error 0 @Extended 0 Return Integer = Success. $iVertOrient is set to Null, returning current Cell Vertical
;				   +										orientation, see constants below.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only the Vertical Orientation Constants listed below are accepted. If $iVertOrient is set to Null the
;					current setting is returned.
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
	;3 = Vert Orient Bottom, 1 = Vert orient Top

	If ($iVertOrient = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $oCell.VertOrient())
	If Not __LOWriter_IntIsBetween($iVertOrient, $LOW_ORIENT_VERT_NONE, $LOW_ORIENT_VERT_BOTTOM) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oCell.VertOrient = $iVertOrient

	Return ($oCell.VertOrient() = $iVertOrient) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)
EndFunc   ;==>_LOWriter_CellVertOrient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateFormatKeyCreate
; Description ...: Create a Date/Time Format Key.
; Syntax ........: _LOWriter_DateFormatKeyCreate(Byref $oDoc, $sFormat)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sFormat             - a string value. The Date/Time format String to create.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFormat not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to Create or Retrieve the Format key, but failed.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Format Key was successfully created, returning Format Key
;				   +												integer.
;				   @Error 0 @Extended 1 Return Integer = Success. Format Key already existed, returning Format Key integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:_LOWriter_DateFormatKeyDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateFormatKeyCreate(ByRef $oDoc, $sFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFormatKey
	Local $tLocale
	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	$iFormatKey = $oFormats.queryKey($sFormat, $tLocale, False)
	If ($iFormatKey > -1) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $iFormatKey) ;Format already existed
	$iFormatKey = $oFormats.addNew($sFormat, $tLocale)
	If ($iFormatKey > -1) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $iFormatKey) ;Format created

	Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ;Failed to create or retrieve Format
EndFunc   ;==>_LOWriter_DateFormatKeyCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateFormatKeyDelete
; Description ...: Delete a User-Created Date/Time Format Key from a Document.
; Syntax ........: _LOWriter_DateFormatKeyDelete(Byref $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $iFormatKey          - an integer value. The User-Created Date/Time format Key to delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = Format Key called in $iFormatKey not found in Document.
;				   @Error 1 @Extended 4 Return 0 = Format Key called in $iFormatKey not User-Created.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to delete key, but Key is still found in Document.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Format Key was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DateFormatKeyCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateFormatKeyDelete(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tLocale
	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_DateFormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ;Key not found.
	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	If ($oFormats.getbykey($iFormatKey).UserDefined() = False) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0) ;Key not User Created.

	$oFormats.removeByKey($iFormatKey)

	Return (_LOWriter_DateFormatKeyExists($oDoc, $iFormatKey) = False) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>_LOWriter_DateFormatKeyDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateFormatKeyExists
; Description ...: Check if a Document contains a Date/Time Format Key Already or not.
; Syntax ........: _LOWriter_DateFormatKeyExists(Byref $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $iFormatKey          - an integer value. The Date Format Key to check for.
; Return values .:  Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iFormatType Parameter for internal Function not an Integer. UDF needs fixed.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to obtain Array of Date/Time Formats.
;				   --Success--
;				   @Error 0 @Extended 0 Return True = Success. Date/Time Format already exists in document.
;				   @Error 0 @Extended 0 Return False = Success. Date/Time Format does not exist in document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateFormatKeyExists(ByRef $oDoc, $iFormatKey)
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = _LOWriter_FormatKeyExists($oDoc, $iFormatKey, $LOW_FORMAT_KEYS_DATE_TIME)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DateFormatKeyExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateFormatKeyGetString
; Description ...: Retrieve a Date/Time Format Key String.
; Syntax ........: _LOWriter_DateFormatKeyGetString(Byref $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $iFormatKey          - an integer value. The Date/Time Format Key to retrieve the string for.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iFormatKey not found in Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve requested Format Key Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returning Format Key's Format String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:_LOWriter_DateFormatKeyList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateFormatKeyGetString(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormatKey

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_DateFormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oFormatKey = $oDoc.getNumberFormats().getByKey($iFormatKey)
	If Not IsObj($oFormatKey) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0) ;Failed to retrieve Key

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oFormatKey.FormatString())
EndFunc   ;==>_LOWriter_DateFormatKeyGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateFormatKeyList
; Description ...: Retrieve an Array of Date/Time Format Keys.
; Syntax ........: _LOWriter_DateFormatKeyList(Byref $oDoc[, $bIsUser = False[, $bUserOnly = False[, $bDateOnly = False[, $bTimeOnly = False]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $bIsUser             - [optional] a boolean value. Default is False. If True, Adds a third column to the
;				   +						return Array with a boolean, whether each Key is user-created or not.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True, only user-created Date/Time
;				   +						Format Keys are returned.
;                  $bDateOnly           - [optional] a boolean value. Default is False. If True, Only Date  FormatKeys are
;				   +						returned.
;                  $bTimeOnly           - [optional] a boolean value. Default is False. If True, Only Time Format Keys are
;				   +						returned.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsUser not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bUserOnly not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bDateOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bTimeOnly not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = Both $bDateOnly and $bTimeOnly set to True. Set one or both to false.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to obtain Array of Date/Time Formats.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning a 2 or three column Array, depending on current
;				   +										$bIsUser setting. Column One (Array[0][0]) will contain the Format
;				   +										Key integer, Column two (Array[0][1]) will contain the Format String
;				   +										And if $bIsUser is set to True, Column Three (Array[0][2]) will
;				   +										contain a Boolean, True if the Format Key is User creater, else
;				   +										false. @Extended is set to the number of Keys returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyDelete, _LOWriter_DateFormatKeyGetString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateFormatKeyList(ByRef $oDoc, $bIsUser = False, $bUserOnly = False, $bDateOnly = False, $bTimeOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys
	Local $avDTFormats[0][3]
	Local $tLocale
	Local $iColumns = 3, $iCount = 0, $iQueryType = $LOW_FORMAT_KEYS_DATE_TIME

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bIsUser) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDateOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bTimeOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($bDateOnly = True) And ($bTimeOnly = True) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$iColumns = ($bIsUser = True) ? $iColumns : 2

	$iQueryType = ($bDateOnly = True) ? $LOW_FORMAT_KEYS_DATE : $iQueryType
	$iQueryType = ($bTimeOnly = True) ? $LOW_FORMAT_KEYS_TIME : $iQueryType

	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	$aiFormatKeys = $oFormats.queryKeys($iQueryType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	ReDim $avDTFormats[UBound($aiFormatKeys)][$iColumns]

	For $i = 0 To UBound($aiFormatKeys) - 1

		If ($bUserOnly = True) Then
			If ($oFormats.getbykey($aiFormatKeys[$i]).UserDefined() = True) Then
				$avDTFormats[$iCount][0] = $aiFormatKeys[$i]
				$avDTFormats[$iCount][1] = $oFormats.getbykey($aiFormatKeys[$i]).FormatString()
				If ($bIsUser = True) Then $avDTFormats[$iCount][2] = $oFormats.getbykey($aiFormatKeys[$i]).UserDefined()
				$iCount += 1
			EndIf
		Else
			$avDTFormats[$i][0] = $aiFormatKeys[$i]
			$avDTFormats[$i][1] = $oFormats.getbykey($aiFormatKeys[$i]).FormatString()
			If ($bIsUser = True) Then $avDTFormats[$i][2] = $oFormats.getbykey($aiFormatKeys[$i]).UserDefined()
		EndIf
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
	Next

	If ($bUserOnly = True) Then ReDim $avDTFormats[$iCount][$iColumns]

	Return SetError($__LOW_STATUS_SUCCESS, UBound($avDTFormats), $avDTFormats)
EndFunc   ;==>_LOWriter_DateFormatKeyList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateStructCreate
; Description ...: Create a Date Structure for inserting a Date into certain other functions.
; Syntax ........: _LOWriter_DateStructCreate([$iYear = Null[, $iMonth = Null[, $iDay = Null[, $iHours = Null[, $iMinutes = Null[, $iSeconds = Null[, $iNanoSeconds = Null[, $bIsUTC = Null]]]]]]]])
; Parameters ....: $iYear               - [optional] an integer value. Default is Null. The Year, in 4 digit integer format.
;                  $iMonth              - [optional] an integer value. Default is Null. The Month, in 2 digit integer format. Set
;				   +						to 0 for Void date. Min 0, Max 12.
;                  $iDay                - [optional] an integer value. Default is Null. The Day, in 2 digit integer format. Set
;				   +						 to 0 for Void date. Min 0, Max 31.
;                  $iHours              - [optional] an integer value. Default is Null. The Hour, in 2 digit integer format. Min
;				   +						0, Max 23.
;                  $iMinutes            - [optional] an integer value. Default is Null. Minutes, in 2 digit integer format. Min
;				   +						0, Max 59.
;                  $iSeconds            - [optional] an integer value. Default is Null. Seconds, in 2 digit integer format. Min
;				   +						0, Max 59.
;                  $iNanoSeconds        - [optional] an integer value. Default is Null. Nano-Second, in integer format. Min 0,
;				   +						Max 999,999,999.
;                  $bIsUTC              - [optional] a boolean value. Default is Null. If true: time zone is UTC Else false:
;				   +						unknown time zone. Libre Office version 4.1 and up.
; Return values .: Success: Structure.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $iYear not an Integer.
;				   @Error 1 @Extended 2 Return 0 = $iYear not 4 digits long.
;				   @Error 1 @Extended 3 Return 0 = $iMonth not an Integer, less than 0 or greater than 12.
;				   @Error 1 @Extended 4 Return 0 = $iDay not an Integer, less than 0 or greater than 31.
;				   @Error 1 @Extended 5 Return 0 = $iHours not an Integer, less than 0 or greater than 23.
;				   @Error 1 @Extended 6 Return 0 = $iMinutes not an Integer, less than 0 or greater than 59.
;				   @Error 1 @Extended 7 Return 0 = $iSeconds not an Integer, less than 0 or greater than 59.
;				   @Error 1 @Extended 8 Return 0 = $iNanoSeconds not an Integer, less than 0 or greater than 999999999.
;				   @Error 1 @Extended 9 Return 0 = $bIsUTC not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.util.DateTime" Object.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.1.
;				   --Success--
;				   @Error 0 @Extended 0 Return Structure = Success. Successfully created the Date/Time Structure,
;				   +												Returning the Date/Time Structure Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateStructCreate($iYear = Null, $iMonth = Null, $iDay = Null, $iHours = Null, $iMinutes = Null, $iSeconds = Null, $iNanoSeconds = Null, $bIsUTC = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tDateStruct

	$tDateStruct = __LOWriter_CreateStruct("com.sun.star.util.DateTime")
	If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($iYear <> Null) Then
		If Not IsInt($iYear) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
		If Not (StringLen($iYear) = 4) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$tDateStruct.Year = $iYear
	Else
		$tDateStruct.Year = @YEAR
	EndIf

	If ($iMonth <> Null) Then
		If Not __LOWriter_IntIsBetween($iMonth, 0, 12) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$tDateStruct.Month = $iMonth
	Else
		$tDateStruct.Month = @MON
	EndIf

	If ($iDay <> Null) Then
		If Not __LOWriter_IntIsBetween($iDay, 0, 31) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$tDateStruct.Day = $iDay
	Else
		$tDateStruct.Day = @MDAY
	EndIf

	If ($iHours <> Null) Then
		If Not __LOWriter_IntIsBetween($iHours, 0, 23) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$tDateStruct.Hours = $iHours
	Else
		$tDateStruct.Hours = @HOUR
	EndIf

	If ($iMinutes <> Null) Then
		If Not __LOWriter_IntIsBetween($iMinutes, 0, 59) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$tDateStruct.Minutes = $iMinutes
	Else
		$tDateStruct.Minutes = @MIN
	EndIf

	If ($iSeconds <> Null) Then
		If Not __LOWriter_IntIsBetween($iSeconds, 0, 59) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tDateStruct.Seconds = $iSeconds
	Else
		$tDateStruct.Seconds = @SEC
	EndIf

	If ($iNanoSeconds <> Null) Then
		If Not __LOWriter_IntIsBetween($iNanoSeconds, 0, 999999999) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$tDateStruct.NanoSeconds = $iNanoSeconds
	Else
		$tDateStruct.NanoSeconds = 0
	EndIf

	If ($bIsUTC <> Null) Then
		If Not IsBool($bIsUTC) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		If Not __LOWriter_VersionCheck(4.1) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$tDateStruct.IsUTC = $bIsUTC
	Else
		If __LOWriter_VersionCheck(4.1) Then $tDateStruct.IsUTC = False
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, $tDateStruct)
EndFunc   ;==>_LOWriter_DateStructCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateStructModify
; Description ...: Set or retrieve Date Structure settings.
; Syntax ........: _LOWriter_DateStructModify(Byref $tDateStruct[, $iYear = Null[, $iMonth = Null[, $iDay = Null[, $iHours = Null[, $iMinutes = Null[, $iSeconds = Null[, $iNanoSeconds = Null[, $bIsUTC = Null]]]]]]]])
; Parameters ....: $tDateStruct         - [in/out] a dll struct value. The Date Structure to modify, returned from a Create or
;				   +						setting retrieval function. Structure will be directly modified.
;                  $iYear               - [optional] an integer value. Default is Null. The Year, in 4 digit integer format.
;                  $iMonth              - [optional] an integer value. Default is Null. The Month, in 2 digit integer format. Set
;				   +						to 0 for Void date. Min 0, Max 12.
;                  $iDay                - [optional] an integer value. Default is Null. The Day, in 2 digit integer format. Set
;				   +						 to 0 for Void date. Min 0, Max 31.
;                  $iHours              - [optional] an integer value. Default is Null. The Hour, in 2 digit integer format. Min
;				   +						0, Max 23.
;                  $iMinutes            - [optional] an integer value. Default is Null. Minutes, in 2 digit integer format. Min
;				   +						0, Max 59.
;                  $iSeconds            - [optional] an integer value. Default is Null. Seconds, in 2 digit integer format. Min
;				   +						0, Max 59.
;                  $iNanoSeconds        - [optional] an integer value. Default is Null. Nano-Second, in integer format. Min 0,
;				   +						Max 999,999,999.
;                  $bIsUTC              - [optional] a boolean value. Default is Null. If true: time zone is UTC Else false:
;				   +						unknown time zone. Libre Office version 4.1 and up.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $tDateStruct not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iYear not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iYear not 4 digits long.
;				   @Error 1 @Extended 4 Return 0 = $iMonth not an Integer, less than 0 or greater than 12.
;				   @Error 1 @Extended 5 Return 0 = $iDay not an Integer, less than 0 or greater than 31.
;				   @Error 1 @Extended 6 Return 0 = $iHours not an Integer, less than 0 or greater than 23.
;				   @Error 1 @Extended 7 Return 0 = $iMinutes not an Integer, less than 0 or greater than 59.
;				   @Error 1 @Extended 8 Return 0 = $iSeconds not an Integer, less than 0 or greater than 59.
;				   @Error 1 @Extended 9 Return 0 = $iNanoSeconds not an Integer, less than 0 or greater than 999999999.
;				   @Error 1 @Extended 10 Return 0 = $bIsUTC not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16, 32, 64, 128
;				   |								1 = Error setting $iYear
;				   |								2 = Error setting $iMonth
;				   |								4 = Error setting $iDay
;				   |								8 = Error setting $iHours
;				   |								16 = Error setting $iMinutes
;				   |								32 = Error setting $iSeconds
;				   |								64 = Error setting $iNanoSeconds
;				   |								128 = Error setting $bIsUTC
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.1.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 7 or 8 Element Array with values in order of function
;				   +								parameters. If current Libre Office version is less than 4.1, the Array
;				   +								will contain 7 elements, as $bIsUTC will be eliminated.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateStructModify(ByRef $tDateStruct, $iYear = Null, $iMonth = Null, $iDay = Null, $iHours = Null, $iMinutes = Null, $iSeconds = Null, $iNanoSeconds = Null, $bIsUTC = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avMod[7]

	If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iYear, $iMonth, $iDay, $iHours, $iMinutes, $iSeconds, $iNanoSeconds, $bIsUTC) Then
		If __LOWriter_VersionCheck(4.1) Then
			__LOWriter_ArrayFill($avMod, $tDateStruct.Year(), $tDateStruct.Month(), $tDateStruct.Day(), $tDateStruct.Hours(), _
					$tDateStruct.Minutes(), $tDateStruct.Seconds(), $tDateStruct.NanoSeconds(), $tDateStruct.IsUTC())
		Else
			__LOWriter_ArrayFill($avMod, $tDateStruct.Year(), $tDateStruct.Month(), $tDateStruct.Day(), $tDateStruct.Hours(), _
					$tDateStruct.Minutes(), $tDateStruct.Seconds(), $tDateStruct.NanoSeconds())
		EndIf

		Return SetError($__LOW_STATUS_SUCCESS, 1, $avMod)
	EndIf

	If ($iYear <> Null) Then
		If Not IsInt($iYear) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		If Not (StringLen($iYear) = 4) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$tDateStruct.Year = $iYear
		$iError = ($tDateStruct.Year() = $iYear) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iMonth <> Null) Then
		If Not __LOWriter_IntIsBetween($iMonth, 0, 12) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$tDateStruct.Month = $iMonth
		$iError = ($tDateStruct.Month() = $iMonth) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iDay <> Null) Then
		If Not __LOWriter_IntIsBetween($iDay, 0, 31) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$tDateStruct.Day = $iDay
		$iError = ($tDateStruct.Day() = $iDay) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iHours <> Null) Then
		If Not __LOWriter_IntIsBetween($iHours, 0, 23) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$tDateStruct.Hours = $iHours
		$iError = ($tDateStruct.Hours() = $iHours) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iMinutes <> Null) Then
		If Not __LOWriter_IntIsBetween($iMinutes, 0, 59) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tDateStruct.Minutes = $iMinutes
		$iError = ($tDateStruct.Minutes() = $iMinutes) ? $iError : BitOR($iError, 16)
	EndIf

	If ($iSeconds <> Null) Then
		If Not __LOWriter_IntIsBetween($iSeconds, 0, 59) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$tDateStruct.Seconds = $iSeconds
		$iError = ($tDateStruct.Seconds() = $iSeconds) ? $iError : BitOR($iError, 32)
	EndIf

	If ($iNanoSeconds <> Null) Then
		If Not __LOWriter_IntIsBetween($iNanoSeconds, 0, 999999999) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$tDateStruct.NanoSeconds = $iNanoSeconds
		$iError = ($tDateStruct.NanoSeconds() = $iNanoSeconds) ? $iError : BitOR($iError, 64)
	EndIf

	If ($bIsUTC <> Null) Then
		If Not IsBool($bIsUTC) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
		If Not __LOWriter_VersionCheck(4.1) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$tDateStruct.IsUTC = $bIsUTC
		$iError = ($tDateStruct.IsUTC() = $bIsUTC) ? $iError : BitOR($iError, 128)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DateStructModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharBorderColor
; Description ...: Set and Retrieve the Character Style Border Line Color by Direct Formatting. Libre Office 4.2 and Up.
; Syntax ........: _LOWriter_DirFrmtCharBorderColor(Byref $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Color of the
;				   +						Character Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value may be used.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Color of the
;				   +						Character Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value may be used.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Color of the
;				   +						Character Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value may be used.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Color of the
;				   +						Character Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value may be used.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						border, Width, Style and Color.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0 or higher than 16,777,215 or not
;				   +								set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0 or higher than 16,777,215 or
;				   +								not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0 or higher than 16,777,215 or not
;				   +								set to Null.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0 or higher than 16,777,215 or
;				   +								not set to Null.
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
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.2.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Border Width must be set first to be able to set Border Style and Color.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				  Call any optional parameter with Null keyword to skip it.
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
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong, _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet, _LOWriter_DirFrmtCharBorderWidth, _LOWriter_DirFrmtCharBorderStyle,
;					_LOWriter_DirFrmtCharBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharBorderColor(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("CharTopBorder")
		$oSelection.setPropertyToDefault("CharBottomBorder") ;Resetting one truly resets all, but just to be sure, reset all.
		$oSelection.setPropertyToDefault("CharLeftBorder")
		$oSelection.setPropertyToDefault("CharRightBorder")
		If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_CharBorder($oSelection, False, False, True, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharBorderPadding
; Description ...: Set and retrieve the distance between the border and the characters by Direct Format. LibreOffice 4.2 and Up.
; Syntax ........: _LOWriter_DirFrmtCharBorderPadding(Byref $oSelection[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four values to the same value.
;				   +											When used, all other parameters are ignored.  In MicroMeters.
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top border distance in
;				   +							MicroMeters.
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom border distance in
;				   +							MicroMeters.
;                  $iLeft               - [optional] an integer value. Default is Null. Set the left border distance in
;				   +							MicroMeters.
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right border distance in
;				   +							MicroMeters.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						border padding, on all sides.
; Return values .:  Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iAll not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $Left not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iRight not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16
;				   |								1 = Error setting $iAll border distance
;				   |								2 = Error setting $iTop border distance
;				   |								4 = Error setting $iBottom border distance
;				   |								8 = Error setting $iLeft border distance
;				   |								16 = Error setting $iRight border distance
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.2.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet, _LOWriter_DirFrmtCharBorderWidth, _LOWriter_DirFrmtCharBorderStyle,
;					_LOWriter_DirFrmtCharBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharBorderPadding(ByRef $oSelection, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		;Resetting any one of these settings causes all to reset; reset the "All" setting for quickness.
		$oSelection.setPropertyToDefault("CharBorderDistance")
		If __LOWriter_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharBorderPadding($oSelection, $iAll, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharBorderStyle
; Description ...: Set or Retrieve the Character Style Border Line style by Direct Format. Libre Office 4.2 and Up.
; Syntax ........: _LOWriter_DirFrmtCharBorderStyle(Byref $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Style of the
;				   +							Character Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Style of the
;				   +							Character Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Style of the
;				   +							Character Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Style of the
;				   +							Character Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						border, Width, Style and Color.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to higher than 17 and not equal to 0x7FFF,
;				   +									Or $iTop is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to higher than 17 and not equal to
;				   +								0x7FFF, Or $iBottom is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to higher than 17 and not equal to 0x7FFF,
;				   +									Or $iLeft is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to higher than 17 and not equal to
;				   +									0x7FFF, Or $iRight is set to less than 0 or not set to Null.
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
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.2.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Border Width must be set first to be able to set Border Style and Color.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Style Constants: $LOW_BORDERSTYLE_NONE(0x7FFF) No border line,
;					$LOW_BORDERSTYLE_SOLID(0) Solid border line,
;					$LOW_BORDERSTYLE_DOTTED(1) Dotted border line,
;					$LOW_BORDERSTYLE_DASHED(2) Dashed border line,
;					$LOW_BORDERSTYLE_DOUBLE(3) Double border line,
;					$LOW_BORDERSTYLE_THINTHICK_SMALLGAP(4) Double border line with a thin line outside and a thick line inside
;						separated by a small gap,
;					$LOW_BORDERSTYLE_THINTHICK_MEDIUMGAP(5) Double border line with a thin line outside and a thick line inside
;						separated by a medium gap,
;						$LOW_BORDERSTYLE_THINTHICK_LARGEGAP(6) Double border line with a thin line outside and a thick line
;						inside separated by a large gap,
;					$LOW_BORDERSTYLE_THICKTHIN_SMALLGAP(7) Double border line with a thick line outside and a thin line inside
;						separated by a small gap,
;					$LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP(8) Double border line with a thick line outside and a thin line inside
;						separated by a medium gap,
;					$LOW_BORDERSTYLE_THICKTHIN_LARGEGAP(9) Double border line with a thick line outside and a thin line inside
;						separated by a large gap,
;					$LOW_BORDERSTYLE_EMBOSSED(10) 3D embossed border line,
;					$LOW_BORDERSTYLE_ENGRAVED(11) 3D engraved border line,
;					$LOW_BORDERSTYLE_OUTSET(12) Outset border line,
;					$LOW_BORDERSTYLE_INSET(13) Inset border line,
;					$LOW_BORDERSTYLE_FINE_DASHED(14) Finely dashed border line,
;					$LOW_BORDERSTYLE_DOUBLE_THIN(15) Double border line consisting of two fixed thin lines separated by a
;						variable gap,
;					$LOW_BORDERSTYLE_DASH_DOT(16) Line consisting of a repetition of one dash and one dot,
;					$LOW_BORDERSTYLE_DASH_DOT_DOT(17) Line consisting of a repetition of one dash and 2 dots.
; Related .......: _LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet, _LOWriter_DirFrmtCharBorderWidth,
;					_LOWriter_DirFrmtCharBorderColor, _LOWriter_DirFrmtCharBorderPadding
;
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharBorderStyle(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("CharTopBorder")
		$oSelection.setPropertyToDefault("CharBottomBorder") ;Resetting one truly resets all, but just to be sure, reset all.
		$oSelection.setPropertyToDefault("CharLeftBorder")
		$oSelection.setPropertyToDefault("CharRightBorder")
		If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_CharBorder($oSelection, False, True, False, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharBorderWidth
; Description ...: Set and Retrieve the Character Style Border Line Width by Direct Formatting. Libre Office 4.2 and Up.
; Syntax ........: _LOWriter_DirFrmtCharBorderWidth(Byref $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line width of the
;				   +							Character Style in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Width of the
;				   +							Character Style in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line width of the
;				   +							Character Style in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Width of the
;				   +							Character Style in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						border, Width, Style and Color.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0 or not set to Null.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.2.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   To "Turn Off" Borders, set them to 0
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Width Constants:  $LOW_BORDERWIDTH_HAIRLINE(2),
;					$LOW_BORDERWIDTH_VERY_THIN(18),
;					$LOW_BORDERWIDTH_THIN(26),
;					$LOW_BORDERWIDTH_MEDIUM(53),
;					$LOW_BORDERWIDTH_THICK(79),
;					$LOW_BORDERWIDTH_EXTRA_THICK(159)
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet, _LOWriter_DirFrmtCharBorderStyle, _LOWriter_DirFrmtCharBorderColor,
;					_LOWriter_DirFrmtCharBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharBorderWidth(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("CharTopBorder")
		$oSelection.setPropertyToDefault("CharBottomBorder") ;Resetting one truly resets all, but just to be sure, reset all.
		$oSelection.setPropertyToDefault("CharLeftBorder")
		$oSelection.setPropertyToDefault("CharRightBorder")
		If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, 0, $iTop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, 0, $iBottom) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, 0, $iLeft) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, 0, $iRight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_CharBorder($oSelection, True, False, False, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharEffect
; Description ...: Set or Retrieve the Font Effect settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharEffect(Byref $oSelection[, $iRelief = Null[, $iCase = Null[, $bHidden = Null[, $bOutline = Null[, $bShadow = Null]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $iRelief             - [optional] an integer value. Default is Null. The Character Relief style. See Constants
;				   +									below.
;                  $iCase               - [optional] an integer value. Default is Null. The Character Case Style. See Constants
;				   +									below.
;                  $bHidden             - [optional] a boolean value. Default is Null. Whether the Characters are hidden or not.
;                  $bOutline            - [optional] a boolean value. Default is Null. Whether the characters have an outline
;				   +									around the outside.
;                  $bShadow             - [optional] a boolean value. Default is Null. Whether the characters have a shadow.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iRelief not an integer or less than 0 or greater than 2. See Constants.
;				   @Error 1 @Extended 5 Return 0 = $iCase not an integer or less than 0 or greater than 4. See Constants.
;				   @Error 1 @Extended 6 Return 0 = $bHidden not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $bOutline not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $bShadow not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4,8, 16
;				   |								1 = Error setting $iRelief
;				   |								2 = Error setting $iCase
;				   |								4 = Error setting $bHidden
;				   |								8 = Error setting $bOutline
;				   |								16 = Error setting $bShadow
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Call a Parameter with Default keyword to clear direct formatting for that setting.
; Relief Constants: $LOW_RELIEF_NONE(0); no relief is used.
;						$LOW_RELIEF_EMBOSSED(1); the font relief is embossed.
;						$LOW_RELIEF_ENGRAVED(2); the font relief is engraved.
; Case Constants : 	$LOW_CASEMAP_NONE(0); The case of the characters is unchanged.
;						$LOW_CASEMAP_UPPER(1); All characters are put in upper case.
;						$LOW_CASEMAP_LOWER(2); All characters are put in lower case.
;						$LOW_CASEMAP_TITLE(3); The first character of each word is put in upper case.
;						$LOW_CASEMAP_SM_CAPS(4); All characters are put in upper case, but with a smaller font height.
; Related .......: _LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharEffect(ByRef $oSelection, $iRelief = Null, $iCase = Null, $bHidden = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_AnyAreDefault($iRelief, $iCase, $bHidden, $bOutline, $bShadow) Then
		If ($iRelief = Default) Then
			$oSelection.setPropertyToDefault("CharRelief")
			$iRelief = Null
		EndIf

		If ($iCase = Default) Then
			$oSelection.setPropertyToDefault("CharCaseMap")
			$iCase = Null
		EndIf

		If ($bHidden = Default) Then
			$oSelection.setPropertyToDefault("CharHidden")
			$bHidden = Null
		EndIf

		If ($bOutline = Default) Then
			$oSelection.setPropertyToDefault("CharContoured")
			$bOutline = Null
		EndIf

		If ($bShadow = Default) Then
			$oSelection.setPropertyToDefault("CharShadowed")
			$bShadow = Null
		EndIf

		If __LOWriter_VarsAreNull($iRelief, $iCase, $bHidden, $bOutline, $bShadow) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharEffect($oSelection, $iRelief, $iCase, $bHidden, $bOutline, $bShadow)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharPosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharPosition(Byref $oSelection[, $bAutoSuper = Null[, $iSuperScript = Null[, $bAutoSub = Null[, $iSubScript = Null[, $iRelativeSize = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $bAutoSuper          - [optional] a boolean value. Default is Null. Whether to active automatically sizing for
;				   +									SuperScript.
;                  $iSuperScript        - [optional] an integer value. Default is Null. SuperScript percentage value. See
;				   +									Remarks.
;                  $bAutoSub            - [optional] a boolean value. Default is Null. Whether to active automatically sizing for
;				   +									SubScript.
;                  $iSubScript          - [optional] an integer value. Default is Null. SubScript percentage value. See Remarks.
;                  $iRelativeSize       - [optional] an integer value. Default is Null. 1-100 percentage relative to current font
;				   +									size.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						Super/Sub Script settings.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoSuper not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bAutoSub not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iSuperScript not an integer, or less than 0, higher than 100 and Not 14000.
;				   @Error 1 @Extended 7 Return 0 = $iSubScript not an integer, or less than -100, higher than 100 and Not 14000.
;				   @Error 1 @Extended 8 Return 0 = $iRelativeSize not an integer, or less than 1, higher than 100.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $iSuperScript
;				   |								2 = Error setting $iSubScript
;				   |								4 = Error setting $iRelativeSize.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Set either $iSubScript or $iSuperScript to 0 to return it to Normal setting.
;					The way LibreOffice is set up Super/SubScript are set in the same setting, Super is a positive number from
;						1 to 100 (percentage), SubScript is a negative number set to 1 to 100 percentage. For the user's
;						convenience this function accepts both positive and negative numbers for SubScript, if a positive number
;						is called for SubScript, it is automatically set to a negative. Automatic Superscript has a integer
;						value of 14000, Auto SubScript has a integer value of -14000. There is no settable setting of Automatic
;						Super/Sub Script, though one exists, it is read-only in LibreOffice, consequently I have made two
;						separate parameters to be able to determine if the user wants to automatically set SuperScript or
;						SubScript. If you set both Auto SuperScript to True and Auto SubScript to True, or $iSuperScript
;						to an integer and $iSubScript to an integer, Subscript will be set as it is the last in the
;						line to be set in this function, and thus will over-write any SuperScript settings.
; Related .......: _LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					 _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					 _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharPosition(ByRef $oSelection, $bAutoSuper = Null, $iSuperScript = Null, $bAutoSub = Null, $iSubScript = Null, $iRelativeSize = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("CharEscapement")
		If __LOWriter_VarsAreNull($bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharPosition($oSelection, $bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharRotateScale
; Description ...: Set or retrieve the character rotational and Scale settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharRotateScale(Byref $oSelection[, $iRotation = Null[, $iScaleWidth = Null[, $bRotateFitLine = Null]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $iRotation           - [optional] an integer value. Default is Null. Degrees to rotate the text. Accepts
;				   +								only 0, 90, and 270 degrees.
;                  $iScaleWidth         - [optional] an integer value. Default is Null. The percentage to  horizontally
;				   +					stretch or compress the text. Must be above 1. Max 100. 100 is normal sizing.
;                  $bRotateFitLine      - [optional] a boolean value. Default is Null. If True, Stretches or compresses the
;				   +						selected text so that it fits between the line that is above the text and the line
;				   +						that is below the text.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iRotation not an Integer or not equal to 0, 90 or 270 degrees.
;				   @Error 1 @Extended 5 Return 0 = $iScaleWidth not an Integer or less than 1% or greater than 100%.
;				   @Error 1 @Extended 6 Return 0 = $bRotateFitLine not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $iRotation
;				   |								2 = Error setting $iScaleWidth
;				   |								4 = Error setting $bRotateFitLine
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 3 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Call a Parameter with Default keyword to clear direct formatting for that setting.
; Related .......: _LOWriter_DirFrmtClear,_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor,_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharRotateScale(ByRef $oSelection, $iRotation = Null, $iScaleWidth = Null, $bRotateFitLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_AnyAreDefault($iRotation, $iScaleWidth, $bRotateFitLine) Then
		If ($iRotation = Default) Then
			$oSelection.setPropertyToDefault("CharRotation")
			$iRotation = Null
		EndIf

		If ($iScaleWidth = Default) Then
			$oSelection.setPropertyToDefault("CharScaleWidth")
			$iScaleWidth = Null
		EndIf

		If ($bRotateFitLine = Default) Then
			$oSelection.setPropertyToDefault("CharRotationIsFitToLine")
			$bRotateFitLine = Null
		EndIf

		If __LOWriter_VarsAreNull($iRotation, $iScaleWidth, $bRotateFitLine) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If __LOWriter_VarsAreNull($iRotation, $iScaleWidth, $bRotateFitLine) Then
		$vReturn = __LOWriter_CharRotateScale($oSelection, $iRotation, $iScaleWidth, $bRotateFitLine)
		__LOWriter_AddTo1DArray($vReturn, $oSelection.CharRotationIsFitToLine())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $vReturn)
	EndIf

	$vReturn = __LOWriter_CharRotateScale($oSelection, $iRotation, $iScaleWidth, $bRotateFitLine)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharRotateScale

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharShadow
; Description ...: Set and retrieve the Shadow for a Character Style by Direct Formatting. Libre Office 4.2 and Up.
; Syntax ........: _LOWriter_DirFrmtCharShadow(Byref $oSelection[, $iWidth = Null[, $iColor = Null[, $bTransparent = Null[, $iLocation = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $iWidth              - [optional] an integer value. Default is Null. Width of the shadow, set in Micrometers.
;                  $iColor              - [optional] an integer value. Default is Null. Color of the shadow. See Remarks and
;				   +							Constants below.
;                  $bTransparent        - [optional] a boolean value. Default is Null. Whether the shadow is transparent or not.
;                  $iLocation           - [optional] an integer value. Default is Null. Location of the shadow compared to the
;				   +									characters. See Constants listed below.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						Character Shadow, Width, Color and Location.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iColor not an Integer, or less than 0 or greater than 16777215 micrometers.
;				   @Error 1 @Extended 6 Return 0 = $bTransparent not a boolean.
;				   @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, or less than 0 or greater than 4. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Shadow format Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Shadow format Object for Error checking.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $iWidth
;				   |								2 = Error setting $iColor
;				   |								4 = Error setting $bTransparent
;				   |								8 = Error setting $iLocation
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.2.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Note: LibreOffice may adjust the set width +/- 1 Micrometer after setting.
;					Color is set in Long Integer format. You can use one of the below listed constants or a custom one.
;Shadow Location Constants: $LOW_SHADOW_NONE(0) = No shadow.
;							$LOW_SHADOW_TOP_LEFT(1) = Shadow is located along the upper and left sides.
;							$LOW_SHADOW_TOP_RIGHT(2) = Shadow is located along the upper and right sides.
;							$LOW_SHADOW_BOTTOM_LEFT(3) = Shadow is located along the lower and left sides.
;							$LOW_SHADOW_BOTTOM_RIGHT(4) = Shadow is located along the lower and right sides.
;Color Constants:  $LOW_COLOR_BLACK(0),
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
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong,  _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer,  _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharShadow(ByRef $oSelection, $iWidth = Null, $iColor = Null, $bTransparent = Null, $iLocation = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("CharShadowFormat")
		If __LOWriter_VarsAreNull($iWidth, $iColor, $bTransparent, $iLocation) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharShadow($oSelection, $iWidth, $iColor, $bTransparent, $iLocation)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtCharSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning)by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtCharSpacing(Byref $oSelection[, $bAutoKerning = Null[, $nKerning = Null]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $bAutoKerning        - [optional] a boolean value. Default is Null. True applies a spacing in between
;				   +						certain pairs of characters. False = disabled.
;                  $nKerning            - [optional] a general number value. Default is Null. The kerning value of the
;				   +								characters. Min is -2 Pt. Max is 928.8 Pt. See Remarks. Values are in
;				   +								Printer's Points as set in the Libre Office UI.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoKerning not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $nKerning not a number, or less than -2 or greater than 928.8 Points.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bAutoKerning
;				   |								2 = Error setting $nKerning.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Call a Parameter with Default keyword to clear direct formatting for that setting.
;					When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User
;						Display, however the internal setting is measured in MicroMeters. They will be automatically converted
;						from Points to MicroMeters and back for retrieval of settings.
;						The acceptable values are from -2 Pt to  928.8 Pt. the figures can be directly converted easily,
;						however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative
;						MicroMeters internally from 928.9 up to 1000 Pt (Max setting). For example, 928.8Pt is the last correct
;						value, which equals 32766 uM (MicroMeters), after this LibreOffice reports the following: ;928.9
;						Pt = -32766 uM; 929 Pt = -32763 uM; 929.1 = -32759; 1000 pt = -30258. Attempting to set Libre's kerning
;						value to anything over 32768 uM causes a COM exception, and attempting to set the kerning to any of
;						these negative numbers sets the User viewable kerning value to -2.0 Pt. For these reasons the max
;						settable kerning is -2.0 Pt  to 928.8 Pt.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtCharSpacing(ByRef $oSelection, $bAutoKerning = Null, $nKerning = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_AnyAreDefault($bAutoKerning, $nKerning) Then
		If ($bAutoKerning = Default) Then
			$oSelection.setPropertyToDefault("CharAutoKerning")
			$bAutoKerning = Null
		EndIf

		If ($nKerning = Default) Then
			$oSelection.setPropertyToDefault("CharKerning")
			$nKerning = Null
		EndIf
		If __LOWriter_VarsAreNull($bAutoKerning, $nKerning) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharSpacing($oSelection, $bAutoKerning, $nKerning)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtCharSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtClear
; Description ...: Clear any Direct formatting in a Cursor or Text Object.
; Syntax ........: _LOWriter_DirFrmtClear(Byref $oDoc, Byref $oSelection)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oSelection          - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 4 Return 0 = $oSelection is a Table Cursor, which is not supported.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;				   @Error 2 @Extended 3 Return 0 = Error retrieving Text Object for creating a ViewCursor Backup.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to determine $oSelection's cursor type.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve document's Viewcursor.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve Text Object for the Viewcursor.
;				   @Error 3 @Extended 4 Return 0 = Failed to a cursor at the position of the View cursor.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Direct Formatting was successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function causes the ViewCursor to select the data input in $oSelection, unless $oSelection is a
;						a ViewCursor object. After the formatting has been cleared the ViewCursor is returned to its previous
;						position.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtClear(ByRef $oDoc, ByRef $oSelection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aArray[0]
	Local $oServiceManager, $oDispatcher, $oText, $oViewCursor, $oViewCursorBackup
	Local $iCursorType

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oSelection)
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	Switch $iCursorType

		Case $LOW_CURTYPE_TEXT_CURSOR, $LOW_CURTYPE_PARAGRAPH, $LOW_CURTYPE_TEXT_PORTION

			;Retrieve the ViewCursor.
			$oViewCursor = $oDoc.CurrentController.getViewCursor()
			If Not IsObj($oViewCursor) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)

			;Create a Text cursor at the current viewCursor position to move the Viewcursor back to.
			$oText = __LOWriter_CursorGetText($oDoc, $oViewCursor)
			If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 3, 0)
			If Not IsObj($oText) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)
			$oViewCursorBackup = $oText.createTextCursorByRange($oViewCursor)
			If Not IsObj($oViewCursorBackup) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 4, 0)

			$oViewCursor.gotoRange($oSelection, False)

			$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:ResetAttributes", "", 0, $aArray)

			;Restore the ViewCursor to its previous location.
			$oViewCursor.gotoRange($oViewCursorBackup, False)

		Case $LOW_CURTYPE_VIEW_CURSOR

			$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:ResetAttributes", "", 0, $aArray)
	EndSwitch

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DirFrmtClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtFont
; Description ...: Set and Retrieve the Font Settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtFont(Byref $oDoc, Byref $oSelection[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $sFontName           - [optional] a string value. Default is Null. The Font Name to change to.
;                  $nFontSize           - [optional] a general number value. Default is Null. The new Font size.
;                  $iPosture            - [optional] an integer value. Default is Null. Italic setting. See Constants below. Also
;				   +								see remarks.
;                  $iWeight             - [optional] an integer value. Default is Null. Bold settings see Constants below.
;				   +								Also see remarks.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 3 Return 0 =  $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 4 Return 0 = $sFontName not available in current document.
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $sFontName not a String.
;				   @Error 1 @Extended 7 Return 0 = $nFontSize not a Number.
;				   @Error 1 @Extended 8 Return 0 = $iPosture not an Integer, less than 0 or greater than 5. See Constants.
;				   @Error 1 @Extended 9 Return 0 = $iWeight less than 50 and not 0, or more than 200. See Constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4,8
;				   |								1 = Error setting $sFontName
;				   |								2 = Error setting $nFontSize
;				   |								4 = Error setting $iPosture
;				   |								8 = Error setting $iWeight
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Call a Parameter with Default keyword to clear direct formatting for that setting.
;					Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted,
;					such as oblique, ultra Bold etc. Libre Writer accepts only the predefined weight values, any other values
;					are changed automatically to an acceptable value, which could trigger a settings error.
; Weight Constants : $LOW_WEIGHT_DONT_KNOW(0); The font weight is not specified/known.
;						$LOW_WEIGHT_THIN(50); specifies a 50% font weight.
;						$LOW_WEIGHT_ULTRA_LIGHT(60); specifies a 60% font weight.
;						$LOW_WEIGHT_LIGHT(75); specifies a 75% font weight.
;						$LOW_WEIGHT_SEMI_LIGHT(90); specifies a 90% font weight.
;						$LOW_WEIGHT_NORMAL(100); specifies a normal font weight.
;						$LOW_WEIGHT_SEMI_BOLD(110); specifies a 110% font weight.
;						$LOW_WEIGHT_BOLD(150); specifies a 150% font weight.
;						$LOW_WEIGHT_ULTRA_BOLD(175); specifies a 175% font weight.
;						$LOW_WEIGHT_BLACK(200); specifies a 200% font weight.
; Slant/Posture Constants : $LOW_POSTURE_NONE(0); specifies a font without slant.
;							$LOW_POSTURE_OBLIQUE(1); specifies an oblique font (slant not designed into the font).
;							$LOW_POSTURE_ITALIC(2); specifies an italic font (slant designed into the font).
;							$LOW_POSTURE_DontKnow(3); specifies a font with an unknown slant.
;							$LOW_POSTURE_REV_OBLIQUE(4); specifies a reverse oblique font (slant not designed into the font).
;							$LOW_POSTURE_REV_ITALIC(5); specifies a reverse italic font (slant designed into the font).
; Related .......: _LOWriter_FontsList, _LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor,
;					 _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor,
;					 _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtFont(ByRef $oDoc, ByRef $oSelection, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_AnyAreDefault($sFontName, $nFontSize, $iPosture, $iWeight) Then
		If ($sFontName = Default) Then
			$oSelection.setPropertyToDefault("CharFontName")
			$sFontName = Null
		EndIf

		If ($nFontSize = Default) Then
			$oSelection.setPropertyToDefault("CharHeight")
			$nFontSize = Null
		EndIf

		If ($iPosture = Default) Then
			$oSelection.setPropertyToDefault("CharPosture")
			$iPosture = Null
		EndIf

		If ($iWeight = Default) Then
			$oSelection.setPropertyToDefault("CharWeight")
			$iWeight = Null
		EndIf

		If __LOWriter_VarsAreNull($sFontName, $nFontSize, $iPosture, $iWeight) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($sFontName <> Null) And Not _LOWriter_FontExists($oDoc, $sFontName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_CharFont($oSelection, $sFontName, $nFontSize, $iPosture, $iWeight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtFontColor(Byref $oSelection[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $iFontColor          - [optional] an integer value. Default is Null. the desired Color value in Long Integer
;				   +								format, to make the font, can be one of the constants listed below or a
;				   +								custom value. Set to $LOW_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - [optional] an integer value. Default is Null. Transparency percentage. 0 is not
;				   +								visible, 100 is fully visible. Available for Libre Office 7.0 and up.
;                  $iHighlight          - [optional] an integer value. Default is Null. A Color value in Long Integer format,
;				   +								to highlight the text in, can be one of the constants listed below or a
;				   +								custom value. Set to -1 for No color.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iFontColor not an integer, less than -1 or greater than 16777215.
;				   @Error 1 @Extended 5 Return 0 = $iTransparency not an Integer, or less than 0 or greater than 100%.
;				   @Error 1 @Extended 6 Return 0 = $iHighlight not an integer, less than -1 or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $FontColor
;				   |								2 = Error setting $iTransparency.
;				   |								4 = Error setting $iHighlight
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 7.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 or 3 Element Array with values in order of function
;				   +								parameters. If The current Libre Office version is below 7.0 the returned
;				   +								array will contain 2 elements, because $iTransparency is not available.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				  Call any optional parameter with Null keyword to skip it.
;					Call a Parameter with Default keyword to clear direct formatting for that setting. Note: Font Color and
;						Transparency reset at the same time as the other, e.g., if you reset Font Color, it will reset
;						Transparency.
; Color Constants: $LOW_COLOR_OFF(-1),
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
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong, _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtFontColor(ByRef $oSelection, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_AnyAreDefault($iFontColor, $iTransparency, $iHighlight) Then
		If ($iFontColor = Default) Then
			$oSelection.setPropertyToDefault("CharColor")
			$iFontColor = Null
		EndIf

		If ($iTransparency = Default) Then
			If Not __LOWriter_VersionCheck(7.0) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
			$oSelection.setPropertyToDefault("CharTransparence")
			$iTransparency = Null
		EndIf

		If ($iHighlight = Default) Then
			If __LOWriter_VersionCheck(4.2) Then $oSelection.setPropertyToDefault("CharHighlight")
			$oSelection.setPropertyToDefault("CharBackColor") ;Both may be used? not sure. Both do the same thing, so reset both to make sure.
			$iHighlight = Null
		EndIf

		If __LOWriter_VarsAreNull($iFontColor, $iTransparency, $iHighlight) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharFontColor($oSelection, $iFontColor, $iTransparency, $iHighlight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtGetCurStyles
; Description ...: Retrieve the current Styles set for a selection of text.
; Syntax ........: _LOWriter_DirFrmtGetCurStyles(Byref $oSelection)
; Parameters ....: $oSelection          - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions that has data selected. Or a paragraph or paragraph section.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support Paragraph Properties service.
;				   @Error 1 @Extended 3 Return 0 = $oSelection does not support Character Properties service.
;				   --Success--
;				   @Error 0 @Extended 0 Return Array = Success. Returns a 4 element array in the following order:
;					Paragraph StyleName, Character StyleName, Page StyleName, Numbering StyleName. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Some of the returned style values may be blank if they are not set, particularly Numberingstyle.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtGetCurStyles(ByRef $oSelection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asStyles[4]

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSelection.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSelection.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	__LOWriter_ArrayFill($asStyles, __LOWriter_ParStyleNameToggle($oSelection.ParaStyleName(), True), _
			__LOWriter_CharStyleNameToggle($oSelection.CharStyleName(), True), _
			__LOWriter_PageStyleNameToggle($oSelection.PageStyleName(), True), _
			$oSelection.NumberingStyleName())

	Return SetError($__LOW_STATUS_SUCCESS, 0, $asStyles)
EndFunc   ;==>_LOWriter_DirFrmtGetCurStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtOverLine
; Description ...: Set and retrieve the OverLine settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtOverLine(Byref $oSelection[, $bWordOnly = Null[, $iOverLineStyle = Null[, $bOLHasColor = Null[, $iOLColor = Null]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not Overlined.
;                  $iOverLineStyle      - [optional] an integer value. Default is Null. The style of the Overline line, see
;				   +									constants listed below. See Remarks.
;                  $bOLHasColor         - [optional] a boolean value. Default is Null. Whether the Overline is colored, must
;				   +						be set to true in order to set the Overline color.
;                  $iOLColor            - [optional] an integer value. Default is Null. The color of the Overline, set in Long
;				   +						integer format. Can be one of the constants below or a custom value. Set to
;				   +						$LOW_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iOverLineStyle not an Integer, or less than 0 or greater than 18. Check
;				   +									the Constants list.
;				   @Error 1 @Extended 6 Return 0 = $bOLHasColor not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iOLColor not an Integer, or less than -1 or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $iOverLineStyle
;				   |								4 = Error setting $OLHasColor
;				   |								8 = Error setting $iOLColor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				  Call any optional parameter with Null keyword to skip it.
;					Call a Parameter with Default keyword to clear direct formatting for that setting. Note: Overline style,
;						Color and $bHasColor all reset together.
;					Note: $bOLHasColor must be set to true in order to set the underline color.
; UnderLine line style Constants: $LOW_UNDERLINE_NONE(0),
;									$LOW_UNDERLINE_SINGLE(1),
;									$LOW_UNDERLINE_DOUBLE(2),
;									$LOW_UNDERLINE_DOTTED(3),
;									$LOW_UNDERLINE_DONT_KNOW(4),
;									$LOW_UNDERLINE_DASH(5),
;									$LOW_UNDERLINE_LONG_DASH(6),
;									$LOW_UNDERLINE_DASH_DOT(7),
;									$LOW_UNDERLINE_DASH_DOT_DOT(8),
;									$LOW_UNDERLINE_SML_WAVE(9),
;									$LOW_UNDERLINE_WAVE(10),
;									$LOW_UNDERLINE_DBL_WAVE(11),
;									$LOW_UNDERLINE_BOLD(12),
;									$LOW_UNDERLINE_BOLD_DOTTED(13),
;									$LOW_UNDERLINE_BOLD_DASH(14),
;									$LOW_UNDERLINE_BOLD_LONG_DASH(15),
;									$LOW_UNDERLINE_BOLD_DASH_DOT(16),
;									$LOW_UNDERLINE_BOLD_DASH_DOT_DOT(17),
;									$LOW_UNDERLINE_BOLD_WAVE(18)
; Color Constants: $LOW_COLOR_OFF(-1),
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
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong, _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtOverLine(ByRef $oSelection, $bWordOnly = Null, $iOverLineStyle = Null, $bOLHasColor = Null, $iOLColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_AnyAreDefault($bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor) Then
		If ($bWordOnly = Default) Then
			$oSelection.setPropertyToDefault("CharWordMode")
			$bWordOnly = Null
		EndIf

		If ($iOverLineStyle = Default) Then
			$oSelection.setPropertyToDefault("CharOverline")
			$iOverLineStyle = Null
		EndIf

		If ($bOLHasColor = Default) Then
			$oSelection.setPropertyToDefault("CharOverlineHasColor")
			$bOLHasColor = Null
		EndIf

		If ($iOLColor = Default) Then
			$oSelection.setPropertyToDefault("CharOverlineColor")
			$iOLColor = Null
		EndIf

		If __LOWriter_VarsAreNull($bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharOverLine($oSelection, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)
	Return SetError(@error, @extended, $vReturn)

EndFunc   ;==>_LOWriter_DirFrmtOverLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParAlignment
; Description ...: Set and Retrieve Alignment settings for a paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParAlignment(Byref $oSelection[, $iHorAlign = Null[, $iVertAlign = Null[, $iLastLineAlign = Null[, $bExpandSingleWord = Null[, $bSnapToGrid = Null[, $iTxtDirection = Null]]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iHorAlign           - [optional] an integer value. Default is Null. The Horizontal alignment of the
;				   +						paragraph. See Constants below. See Remarks.
;                  $iVertAlign          - [optional] an integer value. Default is Null. The Vertical alignment of the
;				   +						paragraph. See Constants below.
;                  $iLastLineAlign      - [optional] an integer value. Default is Null. Specify the alignment for the last line
;				   +						in the paragraph. See Constants below. See Remarks.
;                  $bExpandSingleWord   - [optional] a boolean value. Default is Null. If the last line of a justified paragraph
;				   +						consists of one word, the word is stretched to the width of the paragraph.
;                  $bSnapToGrid         - [optional] a boolean value. Default is Null. If True, Aligns the paragraph to a text
;				   +						grid (if one is active).
;                  $iTxtDirection       - [optional] an integer value. Default is Null. The Text Writing Direction. See Constants
;				   +						below. [Libre Office Default is 4]
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iHorAlign not an integer, less than 0 or greater than 3.
;				   @Error 1 @Extended 5 Return 0 = $iVertAlign not an integer, less than 0 or more than 4.
;				   @Error 1 @Extended 6 Return 0 = $iLastLineAlign not an integer, less than 0 or more than 3.
;				   @Error 1 @Extended 7 Return 0 = $bExpandSingleWord not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $bSnapToGrid not a Boolean.
;				   @Error 1 @Extended 9 Return 0 = $iTxtDirection not an Integer, less than 0 or greater than 5, see constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16, 32
;				   |								1 = Error setting $iHorAlign
;				   |								2 = Error setting $iVertAlign
;				   |								4 = Error setting $iLastLineALign
;				   |								8 = Error setting $bExpandSIngleWord
;				   |								16 = Error setting $bSnapToGrid
;				   |								32 = Error setting $iTxtDirection
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 6 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;					Note: $iTxtDirection constants 2,3, and 5 may not be available depending on your language settings.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Call a Parameter with Default keyword to clear direct formatting for that setting. Note: $iHorAlign,
;					$iLastLineAlign, and $bExpandSingleWord are all reset together.
;					 Note: $iHorAlign must be set to $LOW_PAR_ALIGN_HOR_JUSTIFIED(2) before you can set $iLastLineAlign, and
;					$iLastLineAlign must be set to $LOW_PAR_LAST_LINE_JUSTIFIED(2) before $bExpandSingleWord can be set.
; Horizontal Alignment Constants: $LOW_PAR_ALIGN_HOR_LEFT(0); The Paragraph is left-aligned between the borders.
;									$LOW_PAR_ALIGN_HOR_RIGHT(1); The Paragraph is right-aligned between the borders.
;									$LOW_PAR_ALIGN_HOR_JUSTIFIED(2); The Paragraph is adjusted to both borders / stretched.
;									$LOW_PAR_ALIGN_HOR_CENTER(3); The Paragraph is centered between the left and right borders.
; Vertical Alignment Constants: $LOW_PAR_ALIGN_VERT_AUTO(0); In automatic mode, horizontal text is aligned to the baseline. The
;										same applies to text that is rotated 90°. Text that is rotated 270 ° is aligned to the
;										center.
;									$LOW_PAR_ALIGN_VERT_BASELINE(1); The text is aligned to the baseline.
;									$LOW_PAR_ALIGN_VERT_TOP(2); The text is aligned to the top.
;									$LOW_PAR_ALIGN_VERT_CENTER(3); The text is aligned to the center.
;									$LOW_PAR_ALIGN_VERT_BOTTOM(4); The text is aligned to bottom.
; Last Line Alignment Constants: $LOW_PAR_LAST_LINE_START(0); The Paragraph is aligned either to the Left border or the right,
;										depending on the current text direction.
;									$LOW_PAR_LAST_LINE_JUSTIFIED(2); The Paragraph is adjusted to both borders / stretched.
;									$LOW_PAR_LAST_LINE_CENTER(3); The Paragraph is centered between the left and right borders.
; Text Direction Constants: $LOW_TXT_DIR_LR_TB(0), — text within lines is written left-to-right. Lines and blocks are placed
;								top-to-bottom. Typically, this is the writing mode for normal "alphabetic" text.
;							$LOW_TXT_DIR_RL_TB(1), — text within a line are written right-to-left. Lines and blocks are placed
;								top-to-bottom. Typically, this writing mode is used in Arabic and Hebrew text.
;							$LOW_TXT_DIR_TB_RL(2), — text within a line is written top-to-bottom. Lines and blocks are placed
;								right-to-left. Typically, this writing mode is used in Chinese and Japanese text.
;							$LOW_TXT_DIR_TB_LR(3), — text within a line is written top-to-bottom. Lines and blocks are placed
;								left-to-right. Typically, this writing mode is used in Mongolian text.
;							$LOW_TXT_DIR_CONTEXT(4), — obtain actual writing mode from the context of the object.
;							$LOW_TXT_DIR_BT_LR(5), — text within a line is written bottom-to-top. Lines and blocks are placed
;								left-to-right. (LibreOffice 6.3)
; Related .......: _LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					 _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					 _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					 _LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParAlignment(ByRef $oSelection, $iHorAlign = Null, $iVertAlign = Null, $iLastLineAlign = Null, $bExpandSingleWord = Null, $bSnapToGrid = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_AnyAreDefault($iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection) Then
		If ($iHorAlign = Default) Then
			$oSelection.setPropertyToDefault("ParaAdjust")
			$iHorAlign = Null
		EndIf

		If ($iVertAlign = Default) Then
			$oSelection.setPropertyToDefault("ParaVertAlignment")
			$iVertAlign = Null
		EndIf

		If ($iLastLineAlign = Default) Then
			$oSelection.setPropertyToDefault("ParaLastLineAdjust")
			$iLastLineAlign = Null
		EndIf

		If ($bExpandSingleWord = Default) Then
			$oSelection.setPropertyToDefault("ParaExpandSingleWord")
			$bExpandSingleWord = Null
		EndIf

		If ($bSnapToGrid = Default) Then
			$oSelection.setPropertyToDefault("SnapToGrid")
			$bSnapToGrid = Null
		EndIf

		If ($iTxtDirection = Default) Then
			$oSelection.setPropertyToDefault("WritingMode")
			$iTxtDirection = Null
		EndIf

		If __LOWriter_VarsAreNull($iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParAlignment($oSelection, $iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParBackColor
; Description ...: Set or Retrieve background color settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParBackColor(Byref $oSelection[, $iBackColor = Null[, $bBackTransparent = Null[, $bClearDirFrmt = False]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iBackColor          - [optional] an integer value. Default is Null. The color to make the background. Set in
;				   +							Long integer format. Can be one of the below constants or a custom value.
;												Set to $LOW_COLOR_OFF(-1) to turn Background color off.
;                  $bBackTransparent    - [optional] a boolean value. Default is Null. Whether the background color is
;				   +						transparent or not. True = visible.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						Background color.
; Return values .:  Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iBackColor not an integer, less than -1 or greater than 16777215.
;				   @Error 1 @Extended 5 Return 0 = $bBackTransparent not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $iBackColor
;				   |								2 = Error setting $bBackTransparent
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Color Constants: $LOW_COLOR_OFF(-1),
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
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong,  _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParBackColor(ByRef $oSelection, $iBackColor = Null, $bBackTransparent = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("ParaBackColor")
		If __LOWriter_VarsAreNull($iBackColor, $bBackTransparent) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParBackColor($oSelection, $iBackColor, $bBackTransparent)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParBorderColor
; Description ...: Set and Retrieve the Paragraph Style Border Line Color. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_DirFrmtParBorderColor(Byref $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Color of the
;				   +						Paragraph Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value can be used.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Color of the
;				   +						Paragraph Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value can be used.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Color of the
;				   +						Paragraph Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value can be used.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Color of the
;				   +						Paragraph Style in Long Color code format. One of the predefined constants listed
;				   +						below can be used, or a custom value can be used.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						the Paragraph Border, Width, Style and Color.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0 or higher than 16,777,215 or not
;				   +								set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0 or higher than 16,777,215 or
;				   +								not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0 or higher than 16,777,215 or not
;				   +								set to Null.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0 or higher than 16,777,215 or
;				   +								not set to Null.
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
;				   @Error 0 @Extended 0 Return 1  = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					 Border Width must be set first to be able to set Border Style and Color.
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
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong,  _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet, _LOWriter_DirFrmtParBorderWidth, _LOWriter_DirFrmtParBorderStyle,
;					_LOWriter_DirFrmtParBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParBorderColor(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("TopBorder")
		$oSelection.setPropertyToDefault("BottomBorder")
		$oSelection.setPropertyToDefault("LeftBorder")
		$oSelection.setPropertyToDefault("RightBorder")

		If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oSelection, False, False, True, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParBorderPadding
; Description ...: Set or retrieve the Border Padding (spacing between the Paragraph and border) settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParBorderPadding(Byref $oSelection[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one
;				   +						distance in Micrometers (uM).
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Distance between the Border
;				   +						and Paragraph in Micrometers(uM).
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Distance between the
;				   +						Border and Paragraph in Micrometers(uM).
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Distance between the Border
;				   +						and Paragraph in Micrometers(uM).
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Distance between the Border
;				   +						and Paragraph in Micrometers(uM).
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						Border padding related settings.
; Return values .: Integer or Array, see Remarks.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iAll not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $Left not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iRight not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16
;				   |								1 = Error setting $iAll border distance
;				   |								2 = Error setting $iTop border distance
;				   |								4 = Error setting $iBottom border distance
;				   |								8 = Error setting $iLeft border distance
;				   |								16 = Error setting $iRight border distance
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer,  _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet,  _LOWriter_DirFrmtParBorderWidth, _LOWriter_DirFrmtParBorderStyle,
;					_LOWriter_DirFrmtParBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParBorderPadding(ByRef $oSelection, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("BorderDistance")
		If __LOWriter_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParBorderPadding($oSelection, $iAll, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParBorderStyle
; Description ...: Set and retrieve the Paragraph Border Line style by Direct Formatting. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_DirFrmtParBorderStyle(Byref $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line Style of the
;				   +							Paragraph Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Style of the
;				   +							Paragraph Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line Style of the
;				   +							Paragraph Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Style of the
;				   +							Paragraph Style using one of the line style constants, See below for list. To
;				   +							skip a parameter, set it to Null.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						the Paragraph Border, Width, Style and Color.
; Return values .:  Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to higher than 17 and not equal to 0x7FFF,
;				   +									Or $iTop is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to higher than 17 and not equal to
;				   +								0x7FFF, Or $iBottom is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to higher than 17 and not equal to 0x7FFF,
;				   +									Or $iLeft is set to less than 0 or not set to Null.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to higher than 17 and not equal to
;				   +									0x7FFF, Or $iRight is set to less than 0 or not set to Null.
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
;				   @Error 0 @Extended 0 Return 1  = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					 Border Width must be set first to be able to set Border Style and Color.
; Style Constants: $LOW_BORDERSTYLE_NONE(0x7FFF) No border line,
;					$LOW_BORDERSTYLE_SOLID(0) Solid border line,
;					$LOW_BORDERSTYLE_DOTTED(1) Dotted border line,
;					$LOW_BORDERSTYLE_DASHED(2) Dashed border line,
;					$LOW_BORDERSTYLE_DOUBLE(3) Double border line,
;					$LOW_BORDERSTYLE_THINTHICK_SMALLGAP(4) Double border line with a thin line outside and a thick line inside
;						separated by a small gap,
;					$LOW_BORDERSTYLE_THINTHICK_MEDIUMGAP(5) Double border line with a thin line outside and a thick line inside
;						separated by a medium gap,
;					$LOW_BORDERSTYLE_THINTHICK_LARGEGAP(6) Double border line with a thin line outside and a thick line
;						inside separated by a large gap,
;					$LOW_BORDERSTYLE_THICKTHIN_SMALLGAP(7) Double border line with a thick line outside and a thin line inside
;						separated by a small gap,
;					$LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP(8) Double border line with a thick line outside and a thin line inside
;						separated by a medium gap,
;					$LOW_BORDERSTYLE_THICKTHIN_LARGEGAP(9) Double border line with a thick line outside and a thin line inside
;						separated by a large gap,
;					$LOW_BORDERSTYLE_EMBOSSED(10) 3D embossed border line,
;					$LOW_BORDERSTYLE_ENGRAVED(11) 3D engraved border line,
;					$LOW_BORDERSTYLE_OUTSET(12) Outset border line,
;					$LOW_BORDERSTYLE_INSET(13) Inset border line,
;					$LOW_BORDERSTYLE_FINE_DASHED(14) Finely dashed border line,
;					$LOW_BORDERSTYLE_DOUBLE_THIN(15) Double border line consisting of two fixed thin lines separated by a
;						variable gap,
;					$LOW_BORDERSTYLE_DASH_DOT(16) Line consisting of a repetition of one dash and one dot,
;					$LOW_BORDERSTYLE_DASH_DOT_DOT(17) Line consisting of a repetition of one dash and 2 dots.
; Related .......: _LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					 _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet,  _LOWriter_DirFrmtParBorderWidth,
;					_LOWriter_DirFrmtParBorderColor, _LOWriter_DirFrmtParBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParBorderStyle(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("TopBorder")
		$oSelection.setPropertyToDefault("BottomBorder")
		$oSelection.setPropertyToDefault("LeftBorder")
		$oSelection.setPropertyToDefault("RightBorder")

		If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oSelection, False, True, False, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParBorderWidth
; Description ...: Set and retrieve the Paragraph Border Line Width, or the Paragraph Connect Border option by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParBorderWidth(Byref $oSelection[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bConnectBorder = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iTop                - [optional] an integer value. Default is Null. Sets the Top Border Line width of the
;				   +							Paragraph in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null. Libre Office Version 3.4 and Up.
;                  $iBottom             - [optional] an integer value. Default is Null. Sets the Bottom Border Line Width of the
;				   +							Paragraph in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null. Libre Office Version 3.4 and Up.
;                  $iLeft               - [optional] an integer value. Default is Null. Sets the Left Border Line width of the
;				   +							Paragraph in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null. Libre Office Version 3.4 and Up.
;                  $iRight              - [optional] an integer value. Default is Null. Sets the Right Border Line Width of the
;				   +							Paragraph in MicroMeters. One of the predefined constants listed below can
;				   +						be used. To skip a parameter, set it to Null. Libre Office Version 3.4 and Up.
;                  $bConnectBorder      - [optional] a boolean value. Default is Null. Determines if borders set for a paragraph
;				   +						are merged with the next paragraph. Note: Borders are only merged if they are
;				   +						identical. Libre Office Version 3.4 and Up.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						the Paragraph Border, Width, Style and Color. Doesn't clear $bConnectBorder. See
;				   +						Remarks.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0 or not set to Null.
;				   @Error 1 @Extended 8 Return 0 = $bConnectBorder Not a Boolean and Not set to Null.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1  = Success.  Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
;				   @Error 0 @Extended 0 Return 3 = Success. $bConnectBorder parameter was set to Default, and rest of
;				   +								parameters were set to Null. Direct formatting has been successfully
;				   +								cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Call $bConnectBorder Parameter with Default keyword to clear direct formatting for that setting.
;					 To "Turn Off" Borders, set them to 0
; Width Constants:  $LOW_BORDERWIDTH_HAIRLINE(2),
;					$LOW_BORDERWIDTH_VERY_THIN(18),
;					$LOW_BORDERWIDTH_THIN(26),
;					$LOW_BORDERWIDTH_MEDIUM(53),
;					$LOW_BORDERWIDTH_THICK(79),
;					$LOW_BORDERWIDTH_EXTRA_THICK(159)
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer,  _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet, _LOWriter_DirFrmtParBorderStyle, _LOWriter_DirFrmtParBorderColor,
;					_LOWriter_DirFrmtParBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParBorderWidth(ByRef $oSelection, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bConnectBorder = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("TopBorder")
		$oSelection.setPropertyToDefault("BottomBorder")
		$oSelection.setPropertyToDefault("LeftBorder")
		$oSelection.setPropertyToDefault("RightBorder")

		If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight, $bConnectBorder) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($bConnectBorder = Default) Then
		$oSelection.setPropertyToDefault("ParaIsConnectBorder")
		$bConnectBorder = Null
		If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight, $bConnectBorder) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 3)
	EndIf

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, 0, $iTop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, 0, $iBottom) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, 0, $iLeft) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, 0, $iRight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	If ($bConnectBorder <> Null) And Not IsBool($bConnectBorder) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)

	If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight, $bConnectBorder) Then
		$vReturn = __LOWriter_Border($oSelection, True, False, False, $iTop, $iBottom, $iLeft, $iRight)
		__LOWriter_AddTo1DArray($vReturn, $oSelection.ParaIsConnectBorder())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $vReturn)
	ElseIf Not __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		$vReturn = __LOWriter_Border($oSelection, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

		If @error Then Return SetError(@error, @extended, $vReturn)
	EndIf
	If ($bConnectBorder <> Null) Then $oSelection.ParaIsConnectBorder = $bConnectBorder

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DirFrmtParBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParDropCaps
; Description ...: Set or Retrieve DropCaps settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParDropCaps(Byref $oDoc, Byref $oSelection[, $iNumChar = Null[, $iLines = Null[, $iSpcTxt = Null[, $bWholeWord = Null[, $sCharStyle = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oSelection             - [in/out] an object. an object. A Cursor Object returned from any Cursor Object
;				   +						creation or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iNumChar            - [optional] an integer value. Default is Null. The number of characters to make into
;				   +									DropCaps. Min is 0, max is 9.
;                  $iLines              - [optional] an integer value. Default is Null. The number of lines to drop down, min is
;				   +								0, max is 9, cannot be 1.
;                  $iSpcTxt             - [optional] an integer value. Default is Null. The distance between the drop cap and the
;				   +								following text. In Micrometers.
;                  $bWholeWord          - [optional] a boolean value. Default is Null. Whether to DropCap the whole first word.
;				   +									(Nullifys $iNumChars.)
;                  $sCharStyle          - [optional] a string value. Default is Null. The character style to use for the
;				   +							DropCaps. See Remarks.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						DropCaps and related settings.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 4 Return 0 = $sCharStyle not found in current document.
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iNumChar not an integer, less than 0 or greater than 9.
;				   @Error 1 @Extended 7 Return 0 = $iLines not an Integer, less than 0, equal to 1 or greater than 9
;				   @Error 1 @Extended 8 Return 0 = $iSpaceTxt not an Integer, or less than 0.
;				   @Error 1 @Extended 9 Return 0 = $bWholeWord not a Boolean.
;				   @Error 1 @Extended 10 Return 0 = $sCharStyle not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving DropCap Format Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16
;				   |								1 = Error setting $iNumChar
;				   |								2 = Error setting $iLines
;				   |								4 = Error setting $iSpcTxt
;				   |								8 = Error setting $bWholeWord
;				   |								16 = Error setting $sCharStyle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Set $iNumChars, $iLines, $iSpcTxt to 0 to disable DropCaps.
;					I am unable to find a way to set Drop Caps character style to "None" as is available in the User Interface.
;					When it is set to "None" Libre returns a blank string ("") but setting it to a blank string throws a COM
;					error/Exception, even when attempting to set it to Libre's own return value without any in-between
;					variables, in case I was mistaken as to it being a blank string, but this still caused a COM error. So
;					consequently, you cannot set Character Style to "None", but you can still disable Drop Caps as noted above.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParDropCaps(ByRef $oDoc, ByRef $oSelection, $iNumChar = Null, $iLines = Null, $iSpcTxt = Null, $bWholeWord = Null, $sCharStyle = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("DropCapFormat")
		If __LOWriter_VarsAreNull($iNumChar, $iLines, $iSpcTxt, $bWholeWord, $sCharStyle) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($sCharStyle <> Null) And Not _LOWriter_CharStyleExists($oDoc, $sCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_ParDropCaps($oSelection, $iNumChar, $iLines, $iSpcTxt, $bWholeWord, $sCharStyle)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParDropCaps

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParHyphenation
; Description ...: Set or Retrieve Hyphenation settings for a paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParHyphenation(Byref $oSelection[, $bAutoHyphen = Null[, $bHyphenNoCaps = Null[, $iMaxHyphens = Null[, $iMinLeadingChar = Null[, $iMinTrailingChar = Null[, $bClearDirFrmt = False]]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $bAutoHyphen         - [optional] a boolean value. Default is Null. Whether  automatic hyphenation is applied.
;                  $bHyphenNoCaps       - [optional] a boolean value. Default is Null.  Setting to true will disable hyphenation
;				   +						of words written in CAPS for this paragraph. Libre 6.4 and up.
;                  $iMaxHyphens         - [optional] an integer value. Default is Null. The maximum number of consecutive
;				   +						hyphens. Min 0, Max 99.
;                  $iMinLeadingChar     - [optional] an integer value. Default is Null. Specifies the minimum number of
;				   +						characters to remain before the hyphen character (when hyphenation is applied).
;				   +						Min 2, max 9.
;                  $iMinTrailingChar    - [optional] an integer value. Default is Null. Specifies the minimum number of
;				   +						characters to remain after the hyphen character (when hyphenation is applied).
;				   +						Min 2, max 9.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						Hyphenation.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoHyphen not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bHyphenNoCaps not  a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iMaxHyphens not an Integer, less than 0, or greater than 99.
;				   @Error 1 @Extended 7 Return 0 = $iMinLeadingChar not an Integer, less than 2 or greater than 9.
;				   @Error 1 @Extended 8 Return 0 = $iMinTrailingChar not an Integer, less than 2 or greater than 9.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16
;				   |								1 = Error setting $bAutoHyphen
;				   |								2 = Error setting $bHyphenNoCaps
;				   |								4 = Error setting $iMaxHyphens
;				   |								8 = Error setting $iMinLeadingChar
;				   |								16 = Error setting $iMinTrailingChar
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 6.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 or 5 Element Array with values in order of function
;				   +								parameters. If the current Libre Office Version is below 6.4, then the
;				   +								Array returned will contain 4 elements because $bHyphenNoCaps is not
;				   +								available.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					 Note: $bAutoHyphen set to True for the rest of the settings to be activated, but they will be still
;					successfully set regardless.
; Related .......: _LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParHyphenation(ByRef $oSelection, $bAutoHyphen = Null, $bHyphenNoCaps = Null, $iMaxHyphens = Null, $iMinLeadingChar = Null, $iMinTrailingChar = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("ParaIsHyphenation") ;Resetting one resets all.
		If __LOWriter_VarsAreNull($bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParHyphenation($oSelection, $bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParHyphenation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParIndent
; Description ...: Set or Retrieve Indent settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParIndent(Byref $oSelection[, $iBeforeTxt = Null[, $iAfterTxt = Null[, $iFirstLine = Null[, $bAutoFirstLine = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iBeforeTxt          - [optional] an integer value. Default is Null. The amount of space that you want
;				   +						to indent the paragraph from the page margin. If you want the paragraph to extend
;				   +						into the page margin, enter a negative number. Set in MicroMeters(uM) Min. -9998989,
;				   +						Max.17094
;                  $iAfterTxt           - [optional] an integer value. Default is Null. The amount of space that you want to
;				   +						indent the paragraph from the page margin. If you want the paragraph to extend into
;				   +						the page margin, enter a negative number. Set in MicroMeters(uM) Min. -9998989,
;				   +						Max.17094
;                  $iFirstLine          - [optional] an integer value. Default is Null. Indents the first line of a paragraph by
;				   +						the amount that you enter. Set in MicroMeters(uM) Min. -57785, Max.17094.
;                  $bAutoFirstLine      - [optional] a boolean value. Default is Null. Whether the first line should be indented
;				   +						automatically.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						Indent related settings.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iBeforeText not an integer, less than -9998989 or more than 17094 uM.
;				   @Error 1 @Extended 5 Return 0 = $iAfterText not an integer, less than -9998989 or more than 17094 uM.
;				   @Error 1 @Extended 6 Return 0 = $iFirstLine not an integer, less than -57785 or more than 17094 uM.
;				   @Error 1 @Extended 7 Return 0 = $bAutoFirstLine not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $iBeforeTxt
;				   |								2 = Error setting $iAfterTxt
;				   |								4 = Error setting $iFirstLine
;				   |								8 = Error setting $bAutoFirstLine
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					 Note: $iFirstLine Indent cannot be set if $bAutoFirstLine is set to True.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer,  _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParIndent(ByRef $oSelection, $iBeforeTxt = Null, $iAfterTxt = Null, $iFirstLine = Null, $bAutoFirstLine = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("ParaLeftMargin") ;Resetting one resets all -- but just in case reset the rest.
		$oSelection.setPropertyToDefault("ParaRightMargin")
		$oSelection.setPropertyToDefault("ParaFirstLineIndent")
		$oSelection.setPropertyToDefault("ParaIsAutoFirstLineIndent")
		If __LOWriter_VarsAreNull($iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParIndent($oSelection, $iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParIndent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParOutLineAndList
; Description ...: Set and Retrieve the Outline and List settings for a paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParOutLineAndList(Byref $oDoc, Byref $oSelection[, $iOutline = Null[, $sNumStyle = Null[, $bParLineCount = Null[, $iLineCountVal = Null]]]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iOutline            - [optional] an integer value. Default is Null. The Outline Level, see Constants below.
;				   +							Min is 0, max is 10.
;                  $sNumStyle           - [optional] a string value. Default is Null. Specifies the name of the style for the
;				   +							Paragraph numbering. Set to "" for None.
;                  $bParLineCount       - [optional] a boolean value. Default is Null. Whether the paragraph is included in the
;				   +							line numbering.
;                  $iLineCountVal       - [optional] an integer value. Default is Null. The start value for numbering if a new
;				   +							numbering starts at this paragraph. Set to 0 for no line numbering restart.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 4 Return 0 = $sNumStyle not found in current document.
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iOutline not an integer, less than 0 or greater than 10.
;				   @Error 1 @Extended 7 Return 0 = $sNumStyle not a String.
;				   @Error 1 @Extended 8 Return 0 = $bParLineCount not a Boolean.
;				   @Error 1 @Extended 9 Return 0 = $iLineCountVal Not an Integer or less than 0.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $iOutline
;				   |								2 = Error setting $sNumStyle
;				   |								4 = Error setting $bParLineCount
;				   |								8 = Error setting $iLineCountVal
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Call a Parameter with Default keyword to clear direct formatting for that setting. Note: $iOutline,
;					$bParLineCount, and $iLineCountVal all are reset together.
;				   Note: In LibreOffice User Interface (UI), there are two options available when applying direct formatting,
;					"Restart numbering at this paragraph", and "start value", these are too glitchy to make available, I am
;					able to set "Restart numbering at this paragraph" to True, but I cannot set it to false, and I am unable
;					to clear either setting once applied, so for those reasons I am not including it in this UDF.
; Outline Constants :$LOW_OUTLINE_BODY(0); Indicates that the paragraph belongs to the body text.
;					$LOW_OUTLINE_LEVEL_1(1), Indicates that the paragraph belongs to the corresponding outline level.
;					$LOW_OUTLINE_LEVEL_2(2),
;					$LOW_OUTLINE_LEVEL_3(3),
;					$LOW_OUTLINE_LEVEL_4(4),
;					$LOW_OUTLINE_LEVEL_5(5),
;					$LOW_OUTLINE_LEVEL_6(6),
;					$LOW_OUTLINE_LEVEL_7(7),
;					$LOW_OUTLINE_LEVEL_8(8),
;					$LOW_OUTLINE_LEVEL_9(9),
;					$LOW_OUTLINE_LEVEL_10(10)
; Related .......: _LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParOutLineAndList(ByRef $oDoc, ByRef $oSelection, $iOutline = Null, $sNumStyle = Null, $bParLineCount = Null, $iLineCountVal = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_AnyAreDefault($iOutline, $sNumStyle, $bParLineCount, $iLineCountVal) Then
		If ($iOutline = Default) Then
			$oSelection.setPropertyToDefault("OutlineLevel")
			$iOutline = Null
		EndIf

		If ($sNumStyle = Default) Then
			$oSelection.NumberingStyleName = "" ;set to no numbering style first in order to reset successfully.
			$oSelection.setPropertyToDefault("NumberingStyleName")
			$sNumStyle = Null
		EndIf

		If ($bParLineCount = Default) Then
			$oSelection.setPropertyToDefault("ParaLineNumberCount")
			$bParLineCount = Null
		EndIf

		If ($iLineCountVal = Default) Then
			$oSelection.setPropertyToDefault("ParaLineNumberStartValue")
			$iLineCountVal = Null
		EndIf

		If __LOWriter_VarsAreNull($iOutline, $sNumStyle, $bParLineCount, $iLineCountVal) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($sNumStyle <> Null) And ($sNumStyle <> "") And Not _LOWriter_NumStyleExists($oDoc, $sNumStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_ParOutLineAndList($oSelection, $iOutline, $sNumStyle, $bParLineCount, $iLineCountVal)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParOutLineAndList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParPageBreak
; Description ...: Set or Retrieve Page Break Settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParPageBreak(Byref $oDoc, Byref $oSelection[, $iBreakType = Null[, $iPgNumOffSet = Null[, $sPageStyle = Null[, $bClearDirFrmt = False]]]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iBreakType          - [optional] an integer value. Default is Null. The Page Break Type. See Constants below.
;                  $iPgNumOffSet        - [optional] an integer value. Default is Null. If a page break property is set at a
;				   +						paragraph, this property contains the new value for the page number.
;                  $sPageStyle          - [optional] a string value. Default is Null. Creates a page break before the paragraph
;				   +						it belongs to and assigns the value as the name of the new page style to use. Note:
;				   +						If you set this parameter, to remove the page break setting you must set this to "".
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						Page Break, including Type, Number offset and Page Style. See Remarks.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 4 Return 0 = $sPageStyle not found in current document.
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iBreakType not an integer, less than 0 or greater than 6.
;				   @Error 1 @Extended 7 Return 0 = $iPgNumOffSet not an Integer or less than 0.
;				   @Error 1 @Extended 8 Return 0 = $sPageStyle not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $iBreakType
;				   |								2 = Error setting $iPgNumOffSet
;				   |								4 = Error setting $sPageStyle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 3 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Note: Clearing directly formatted page breaks may fail, If the cursor selection contains more than one
;					paragraph that has more than one type of page break, it may fail to literally  reset it to the paragraph
;					style's original settings even though it returns a success, you will need to reset each paragraph one at
;					a time if this is the case.
;					Note: Break Type must be set before PageStyle will be able to be set, and page style needs set before
;					$iPgNumOffSet can be set.
;					Libre doesn't directly show in its User interface options for Break type constants #3 and #6 (Column both)
;						and (Page both), but  doesn't throw an error when being set to either one, so they are included here,
;						 though I'm not sure if they will work correctly.
;Break Constants : $LOW_BREAK_NONE(0) – No column or page break is applied.
;						$LOW_BREAK_COLUMN_BEFORE(1) – A column break is applied before the current Paragraph. The current
;							Paragraph, therefore, is the first in the column.
;						$LOW_BREAK_COLUMN_AFTER(2) – A column break is applied after the current Paragraph. The current
;							Paragraph, therefore, is the last in the column.
;						$LOW_BREAK_COLUMN_BOTH(3) – A column break is applied before and after the current Paragraph. The
;							current Paragraph, therefore, is the only Paragraph in the column.
;						$LOW_BREAK_PAGE_BEFORE(4) – A page break is applied before the current Paragraph. The current Paragraph,
;						therefore, is the first on the page.
;						$LOW_BREAK_PAGE_AFTER(5) – A page break is applied after the current Paragraph. The current Paragraph,
;						therefore, is the last on the page.
;						$LOW_BREAK_PAGE_BOTH(6) – A page break is applied before and after the current Paragraph. The current
;						Paragraph, therefore, is the only paragraph on the page.
; Related .......: _LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParPageBreak(ByRef $oDoc, ByRef $oSelection, $iBreakType = Null, $iPgNumOffSet = Null, $sPageStyle = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If $bClearDirFrmt Then
		$oSelection.PageDescName = ""
		$oSelection.BreakType = $LOW_BREAK_NONE
		$oSelection.setPropertyToDefault("BreakType")
		$oSelection.setPropertyToDefault("PageNumberOffset")
		$oSelection.setPropertyToDefault("PageDescName")

		If __LOWriter_VarsAreNull($iBreakType, $iPgNumOffSet, $sPageStyle) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If ($sPageStyle <> Null) And Not _LOWriter_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_ParPageBreak($oSelection, $iBreakType, $iPgNumOffSet, $sPageStyle)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParPageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParShadow
; Description ...: Set or Retrieve the Shadow settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParShadow(Byref $oSelection[, $iWidth = Null[, $iColor = Null[, $bTransparent = Null[, $iLocation = Null[, $bClearDirFrmt = False]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the shadow set in
;				   +							Micrometers.
;                  $iColor              - [optional] an integer value. Default is Null. The color of the shadow, set in Long
;				   +						Integer format. Can be one of the below constants or a custom value. Set to
;
;                  $bTransparent        - [optional] a boolean value. Default is Null. Whether or not the shadow is transparent.
;                  $iLocation           - [optional] an integer value. Default is Null. The location of the shadow compared to
;				   +								the paragraph. See Constants below.
;                  $bClearDirFrmt       - [optional] a boolean value. Default is False. If True, clears ALL direct formatting of
;				   +						 Shadow Width, Color and Location.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iWidth not an integer or less than 0.
;				   @Error 1 @Extended 5 Return 0 = $iColor not an integer, less than 0 or greater than 16777215.
;				   @Error 1 @Extended 6 Return 0 = $bTransparent not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, less than 0 or greater than 4. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Shadow Format Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Shadow Format Object for Error checking.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4,8
;				   |								1 = Error setting $iWidth
;				   |								2 = Error setting $iColor
;				   |								4 = Error setting $bTransparent
;				   |								8 = Error setting $iLocation
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. $bClearDirFrmt was set to True, and rest of parameters were set
;				   +								to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Note: LibreOffice may change the shadow width +/- a Micrometer.
; Shadow Location Constants: $LOW_SHADOW_NONE(0) = No shadow.
;							$LOW_SHADOW_TOP_LEFT(1) = Shadow is located along the upper and left sides.
;							$LOW_SHADOW_TOP_RIGHT(2) = Shadow is located along the upper and right sides.
;							$LOW_SHADOW_BOTTOM_LEFT(3) = Shadow is located along the lower and left sides.
;							$LOW_SHADOW_BOTTOM_RIGHT(4) = Shadow is located along the lower and right sides.
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
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong,  _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer,  _LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor,
;					_LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor,
;					_LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor,
;					_LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
;
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParShadow(ByRef $oSelection, $iWidth = Null, $iColor = Null, $bTransparent = Null, $iLocation = Null, $bClearDirFrmt = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If $bClearDirFrmt Then
		$oSelection.setPropertyToDefault("ParaShadowFormat")
		If __LOWriter_VarsAreNull($iWidth, $iColor, $bTransparent, $iLocation) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParShadow($oSelection, $iWidth, $iColor, $bTransparent, $iLocation)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParSpace
; Description ...: Set and Retrieve Line Spacing settings for a paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParSpace(Byref $oSelection[, $iAbovePar = Null[, $iBelowPar = Null[, $bAddSpace = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null[, $bPageLineSpc = Null]]]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iAbovePar           - [optional] an integer value. Default is Null. The Space above a paragraph, in
;				   +									Micrometers. Min 0 Micrometers (uM) Max 10,008 uM.
;                  $iBelowPar           - [optional] an integer value. Default is Null. The Space Below a paragraph, in
;				   +									Micrometers. Min 0 Micrometers (uM) Max 10,008 uM.
;                  $bAddSpace           - [optional] a boolean value. Default is Null. If true, the top and bottom margins
;				   +									of the paragraph should not be applied when the previous and next
;				   +									paragraphs have the same style. Libre Office 3.6 and Up.
;                  $iLineSpcMode        - [optional] an integer value. Default is Null. The type of the line spacing of a
;				   +									paragraph. See Constants below, also notice min and max values for each.
;                  $iLineSpcHeight      - [optional] an integer value. Default is Null. This value specifies the spacing
;				   +						of the lines. See Remarks for Minimum and Max values.
;                  $bPageLineSpc        - [optional] a boolean value. Default is Null. Determines if the register mode is
;				   +						applied to a paragraph. See Remarks.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iAbovePar not an integer, less than 0 or more than 10008 uM.
;				   @Error 1 @Extended 5 Return 0 = $iBelowPar not an integer, less than 0 or more than 10008 uM.
;				   @Error 1 @Extended 6 Return 0 = $bAddSpc not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iLineSpcMode Not an integer, less than 0 or greater than 3. See Constants.
;				   @Error 1 @Extended 8 Return 0 = $iLineSpcHeight not an Integer.
;				   @Error 1 @Extended 9 Return 0 = $iLineSpcMode set to 0(Proportional) and $iLineSpcHeight less than 6(%)
;				   +									or greater than 65535(%).
;				   @Error 1 @Extended 10 Return 0 = $iLineSpcMode set to 1 or 2(Minimum, or Leading) and $iLineSpcHeight less
;				   +								than 0 uM or greater than 10008 uM
;				   @Error 1 @Extended 11 Return 0 = $iLineSpcMode set to 3(Fixed) and $iLineSpcHeight less than 51 uM
;				   +									or greater than 10008 uM.
;				   @Error 1 @Extended 12 Return 0 = $bPageLineSpc not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaLineSpacing Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16, 32
;				   |								1 = Error setting $iAbovePar
;				   |								2 = Error setting $iBelowPar
;				   |								4 = Error setting $bAddSpace
;				   |								8 = Error setting $iLineSpcMode
;				   |								16 = Error setting $iLineSpcHeight
;				   |								32 = Error setting bPageLineSpc
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 or 6 Element Array with values in order of function
;				   +								parameters. If the current Libre Office version is less than 3.6, the
;				   +								returned Array will contain 5 elements, because $bAddSpace is not available.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Call a Parameter with Default keyword to clear direct formatting for that setting. Note: $iAbovePar,
;					$iBelowPar, and $bAddSpace are all reset together, $iLineSpace Mode / Height also reset together.
;					Note:  $bPageLineSpc(Register mode) is only used if the register mode property of the page style is switched
;						on. $bPageLineSpc(Register Mode) Aligns the baseline of each line of text to a vertical document grid,
;						so that each line is the same height.
;					Note: The settings in Libre Office, (Single,1.15, 1.5, Double,) Use the Proportional mode, and are just
;						varying percentages. e.g Single = 100, 1.15 = 115%, 1.5 = 150%, Double = 200%.
;					$iLineSpcHeight depends on the $iLineSpcMode used, see constants for accepted Input values.
;					Note: $iAbovePar, $iBelowPar, $iLineSpcHeight may change +/- 1 MicroMeter once set.
; Spacing Constants :$LOW_LINE_SPC_MODE_PROP(0); This specifies the height value as a proportional value. Min 6% Max 65,535%.
;							(without percentage sign)
;						$LOW_LINE_SPC_MODE_MIN(1); (Minimum/At least) This specifies the height as the minimum line height.
;							Min 0, Max 10008 MicroMeters (uM)
;						$LOW_LINE_SPC_MODE_LEADING(2); This specifies the height value as the distance to the previous line.
;							Min 0, Max 10008 MicroMeters (uM)
;						$LOW_LINE_SPC_MODE_FIX(3); This specifies the height value as a fixed line height. Min 51 MicroMeters,
;							Max 10008 MicroMeters (uM)
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParSpace(ByRef $oSelection, $iAbovePar = Null, $iBelowPar = Null, $bAddSpace = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null, $bPageLineSpc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_AnyAreDefault($iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc) Then
		If ($iAbovePar = Default) Then
			$oSelection.setPropertyToDefault("ParaTopMargin")
			$iAbovePar = Null
		EndIf

		If ($iBelowPar = Default) Then
			$oSelection.setPropertyToDefault("ParaBottomMargin")
			$iBelowPar = Null
		EndIf

		If ($bAddSpace = Default) Then
			If Not __LOWriter_VersionCheck(3.6) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
			$oSelection.setPropertyToDefault("ParaContextMargin")
			$bAddSpace = Null
		EndIf

		If ($iLineSpcMode = Default) Or ($iLineSpcHeight = Default) Then
			$oSelection.setPropertyToDefault("ParaLineSpacing")
			$iLineSpcMode = Null
		EndIf

		If ($iLineSpcHeight = Default) Or ($iLineSpcHeight = Default) Then
			$oSelection.setPropertyToDefault("ParaLineSpacing")
			$iLineSpcHeight = Null
		EndIf

		If ($bPageLineSpc = Default) Then
			$oSelection.setPropertyToDefault("ParaRegisterModeActive")
			$bPageLineSpc = Null
		EndIf

		If __LOWriter_VarsAreNull($iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParSpace($oSelection, $iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParSpace

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParTabStopCreate
; Description ...: Create a new TabStop for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParTabStopCreate(Byref $oSelection, $iPosition[, $iFillChar = Null[, $iAlignment = Null[, $iDecChar = Null]]])
; Parameters ....: $oSelection          - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iPosition           - an integer value. The TabStop position/length to set the new TabStop to. Set in
;				   +						Micrometers (uM). See Remarks.
;                  $iAlignment          - [optional] an integer value. Default is Null. The Asc (see autoit function) value of
;				   +						any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
;                  $iFillChar           - [optional] an integer value. Default is Null. The position of where the end of a Tab
;				   +						is aligned to compared to the text. See Constants.
;                  $iDecChar            - [optional] an integer value. Default is Null. Enter a character(in Asc Value(See
;				   +						Autoit Function)) that you want the decimal tab to use as a decimal separator. Can
;				   +						only be set if $iAlignment is set to $LOW_TAB_ALIGN_DECIMAL.
; Return values .:  Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection parameter not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = $iPosition not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iPosition Already exists in this ParStyle.
;				   @Error 1 @Extended 5 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iFillChar not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants.
;				   @Error 1 @Extended 8 Return 0 = $iDecChar not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Array Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.style.TabStop" Object.
;				   @Error 2 @Extended 3 Return 0 = Error retrieving list of TabStop Positions.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to identify the new Tabstop once inserted. in $iPosition.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return Integer = Some settings were not successfully set. Use BitAND to test
;				   +							@Extended for the following values: 1, 2, 4, 8
;				   |								1 = Error setting $iPosition
;				   |								2 = Error setting $iFillChar
;				   |								4 = Error setting $iAlignment
;				   |								8 = Error setting $iDecChar
;				   |						Note: $iNewTabStop position is still returned as even though some settings weren't
;				   +						successfully set, the new TabStop was still created.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Settings were successfully set. New TabStop position
;				   +								is returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
;					Note: $iPosition once set can vary +/- 1 uM. To ensure you can identify the tabstop to modify it again,
;						This function returns the new TabStop position in @Extended when $iPosition is set, return value will
;						be set to 2. See Return Values.
;					Note: Since $iPosition can fluctuate +/- 1 uM when it is inserted into LibreOffice, it is possible to
;						accidentally overwrite an already existing TabStop.
;					Note: $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32.
;						The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can
;						also enter a custom ASC value. See ASC Autoit Func. and "ASCII Character Codes" in the Autoit help file.
;					Call any optional parameter with Null keyword to skip it.
; Tab Alignment Constants: $LOW_TAB_ALIGN_LEFT(0); Aligns the left edge of the text to the tab stop and extends the text to the
;								right.
;							$LOW_TAB_ALIGN_CENTER(1); Aligns the center of the text to the tab stop
;							$LOW_TAB_ALIGN_RIGHT(2); Aligns the right edge of the text to the tab stop and extends the text to
;								the left of the tab stop.
;							$LOW_TAB_ALIGN_DECIMAL(3); Aligns the decimal separator of a number to the center of the tab stop
;								and text to the left of the tab
;							$LOW_TAB_ALIGN_DEFAULT(4);4 = This setting is the default, setting when no TabStops are present.
;								Setting and Tabstop to this constant will make it disappear from the TabStop list. It is
;								therefore only listed here for property reading purposes.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DirFrmtParTabStopDelete,
;					_LOWriter_DirFrmtParTabStopMod, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
;
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParTabStopCreate(ByRef $oSelection, $iPosition, $iFillChar = Null, $iAlignment = Null, $iDecChar = Null)

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If Not IsInt($iPosition) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If __LOWriter_ParHasTabStop($oSelection, $iPosition) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$iPosition = __LOWriter_ParTabStopCreate($oSelection, $iPosition, $iAlignment, $iFillChar, $iDecChar)
	Return SetError(@error, @extended, $iPosition)
EndFunc   ;==>_LOWriter_DirFrmtParTabStopCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParTabStopDelete
; Description ...: Delete a TabStop from a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParTabStopDelete(Byref $oDoc, Byref $oSelection, $iTabStop)
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oSelection      - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iTabStop            - an integer value. The Tab position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 4 Return 0 = $iTabStop not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iTabStop not found in this ParStyle.
;				   @Error 1 @Extended 6 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 7 Return 0 = Passed Document Object to internal function not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to identify and delete TabStop in Paragraph.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Returns true if the TabStop was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin.
;						This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be
;						one of a certain length per document.
; Related .......: _LOWriter_DirFrmtParTabStopCreate, _LOWriter_DirFrmtParTabStopList, _LOWriter_DocGetViewCursor,
;					_LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor,
;					_LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor,
;					_LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
;
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParTabStopDelete(ByRef $oDoc, ByRef $oSelection, $iTabStop)
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iTabStop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not __LOWriter_ParHasTabStop($oSelection, $iTabStop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_ParTabStopDelete($oSelection, $oDoc, $iTabStop)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParTabStopDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParTabStopList
; Description ...: Retrieve a List of TabStops available in a Paragraph from Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParTabStopList(Byref $oSelection)
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
; Return values .:  Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection parameter not an Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. An Array of TabStops. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
; Related .......: _LOWriter_DirFrmtParTabStopCreate, _LOWriter_DirFrmtParTabStopDelete, _LOWriter_DirFrmtParTabStopMod,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParTabStopList(ByRef $oSelection)
	Local $aiTabList

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$aiTabList = __LOWriter_ParTabStopList($oSelection)

	Return SetError(@error, @extended, $aiTabList)
EndFunc   ;==>_LOWriter_DirFrmtParTabStopList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop in a Paragraph from Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParTabStopMod(Byref $oSelection, $iTabStop[, $iPosition = Null[, $iFillChar = Null[, $iAlignment = Null[, $iDecChar = Null]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $iTabStop            - an integer value. The Tab position of the TabStop to modify. See Remarks.
;                  $iPosition           - [optional] an integer value. Default is Null. The New position to set the input
;				   +						position to. Set in Micrometers (uM). See Remarks.
;                  $iFillChar           - [optional] an integer value. Default is Null. The Asc (see autoit function) value of
;				   +						any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
;                  $iAlignment          - [optional] an integer value. Default is Null. The position of where the end of a Tab is
;				   +						aligned to compared to the text. See Constants.
;                  $iDecChar            - [optional] an integer value. Default is Null. Enter a character(in Asc Value(See
;				   +						Autoit Function)) that you want the decimal tab to use as a decimal separator. Can
;				   +						only be set if $iAlignment is set to $LOW_TAB_ALIGN_DECIMAL.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = $iTabStop not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iTabStop not found in this ParStyle.
;				   @Error 1 @Extended 5 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iPosition not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iFillChar not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants.
;				   @Error 1 @Extended 9 Return 0 = $iDecChar not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Requested TabStop Object.
;				   @Error 2 @Extended 3 Return 0 = Error retrieving list of TabStop Positions.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Paragraph style already contains a TabStop at the length/Position specified
;				   +		in $iPosition.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $iPosition
;				   |								2 = Error setting $iFillChar
;				   |								4 = Error setting $iAlignment
;				   |								8 = Error setting $iDecChar
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended ? Return 2 = Success. Settings were successfully set. New TabStop position is returned
;				   +								 in @Extended.
;				   @Error 0 @Extended 0 Return 3 = Success. $iTabStop parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting inserted TabStops have been successfully
;				   +								cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Call a $iTabStop with Default keyword to clear all direct formatting created TabStops.
;					Note: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page
;						margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there
;						can only be one of a certain length per document.
;					Note: $iPosition once set can vary +/- 1 uM. To ensure you can identify the tabstop to modify it again,
;						This function returns the new TabStop position in @Extended when $iPosition is set, return value will
;						be set to 2. See Return Values.
;					Note: Since $iPosition can fluctuate +/- 1 uM when it is inserted into LibreOffice, it is possible to
;						accidentally overwrite an already existing TabStop.
;					Note: $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32.
;						The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can
;						also enter a custom ASC value. See ASC Autoit Func. and "ASCII Character Codes" in the Autoit help file.
; Tab Alignment Constants: $LOW_TAB_ALIGN_LEFT(0); Aligns the left edge of the text to the tab stop and extends the text to the
;								right.
;							$LOW_TAB_ALIGN_CENTER(1); Aligns the center of the text to the tab stop
;							$LOW_TAB_ALIGN_RIGHT(2); Aligns the right edge of the text to the tab stop and extends the text to
;								the left of the tab stop.
;							$LOW_TAB_ALIGN_DECIMAL(3); Aligns the decimal separator of a number to the center of the tab stop
;								and text to the left of the tab
;							$LOW_TAB_ALIGN_DEFAULT(4);4 = This setting is the default, setting when no TabStops are present.
;								Setting and Tabstop to this constant will make it disappear from the TabStop list. It is
;								therefore only listed here for property reading purposes.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DirFrmtParTabStopCreate,
;					_LOWriter_DirFrmtParTabStopDelete, _LOWriter_DirFrmtParTabStopList,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParTabStopMod(ByRef $oSelection, $iTabStop, $iPosition = Null, $iFillChar = Null, $iAlignment = Null, $iDecChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($iTabStop = Default) Then
		$oSelection.setPropertyToDefault("ParaTabStops")
		Return SetError($__LOW_STATUS_SUCCESS, 0, 3)
	EndIf

	If Not IsInt($iTabStop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not __LOWriter_ParHasTabStop($oSelection, $iTabStop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_ParTabStopMod($oSelection, $iTabStop, $iPosition, $iFillChar, $iAlignment, $iDecChar)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParTabStopMod

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtParTxtFlowOpt
; Description ...: Set and Retrieve Text Flow settings for a Paragraph by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtParTxtFlowOpt(Byref $oSelection[, $bParSplit = Null[, $bKeepTogether = Null[, $iParOrphans = Null[, $iParWidows = Null]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval functions, Or A Paragraph Object/Object Section returned from
;				   +						_LOWriter_ParObjCreateList or _LOWriter_ParObjSectionsGet function.
;                  $bParSplit           - [optional] a boolean value. Default is Null.  FALSE prevents the paragraph from
;				   +						getting split into two pages or columns
;                  $bKeepTogether       - [optional] a boolean value. Default is Null. TRUE prevents page or column breaks
;				   +						 between this and the following paragraph
;                  $iParOrphans         - [optional] an integer value. Default is Null. Specifies the minimum number of lines
;				   +							of the paragraph that have to be at bottom of a page if the paragraph is spread
;				   +							over more than one page. Min is 0 (disabled), and cannot be 1. Max is 9.
;                  $iParWidows          - [optional] an integer value. Default is Null. Specifies the minimum number of lines
;				   +						of the paragraph that have to be at top of a page if the paragraph is spread over
;				   +						more than one page. Min is 0 (disabled), and cannot be 1. Max is 9.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection not a Cursor Object and not a Paragraph portion Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bParSplit not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bKeepTogether not  a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iParOrphans not an Integer, less than 0, equal to 1, or greater than 9.
;				   @Error 1 @Extended 7 Return 0 = $iParWidows not an Integer, less than 0, equal to 1, or greater than 9.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $bParSplit
;				   |								2 = Error setting $bKeepTogether
;				   |								4 = Error setting $iParOrphans
;				   |								8 = Error setting $iParWidows
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Call a Parameter with Default keyword to clear direct formatting for that setting. Note: Resetting Orphan or
;					Widow will reset $bParSplit to False if it was set to True.
;					 Note: If you do not set ParSplit to True, the rest of the settings will still show to have been set but
;					will not become active until $bParSplit is set to true.
; Related .......:_LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtParTxtFlowOpt(ByRef $oSelection, $bParSplit = Null, $bKeepTogether = Null, $iParOrphans = Null, $iParWidows = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_AnyAreDefault($bParSplit, $bKeepTogether, $iParOrphans, $iParWidows) Then
		If ($bParSplit = Default) Then
			$oSelection.setPropertyToDefault("ParaSplit")
			$bParSplit = Null
		EndIf

		If ($bKeepTogether = Default) Then
			$oSelection.setPropertyToDefault("ParaKeepTogether")
			$bKeepTogether = Null
		EndIf

		If ($iParOrphans = Default) Then
			$oSelection.setPropertyToDefault("ParaOrphans")
			$iParOrphans = Null
		EndIf

		If ($iParWidows = Default) Then
			$oSelection.setPropertyToDefault("ParaWidows")
			$iParWidows = Null
		EndIf

		If __LOWriter_VarsAreNull($bParSplit, $bKeepTogether, $iParOrphans, $iParWidows) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_ParTxtFlowOpt($oSelection, $bParSplit, $bKeepTogether, $iParOrphans, $iParWidows)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtParTxtFlowOpt

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtStrikeOut
; Description ...: Set or Retrieve the StrikeOut settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtStrikeOut(Byref $oSelection[, $bWordOnly = Null[, $bStrikeOut = Null[, $iStrikeLineStyle = Null]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. Whether to strike out words only and skip
;				   +							whitespaces. True = skip whitespaces.
;                  $bStrikeOut          - [optional] a boolean value. Default is Null. True = strikeout, False = no strikeout.
;                  $iStrikeLineStyle    - [optional] an integer value. Default is Null. The Strikeout Line Style, see constants
;				   +								below.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bStrikeOut not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iStrikeLineStyle not an Integer, or less than 0 or greater than 8. Check
;				   +									the Constants list.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $bStrikeOut
;				   |								4 = Error setting $iStrikeLineStyle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 3 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Call a Parameter with Default keyword to clear direct formatting for that setting. Note: $bStrikeout and
;						$iStrikeLineStyle are reset together.
;					Note Strikeout converted to single line in Ms word document format.
; Strikeout Line Style Constants : $LOW_STRIKEOUT_NONE(0); specifies not to strike out the characters.
;					$LOW_STRIKEOUT_SINGLE(1); specifies to strike out the characters with a single line
;					$LOW_STRIKEOUT_DOUBLE(2); specifies to strike out the characters with a double line.
;					$LOW_STRIKEOUT_DONT_KNOW(3); The strikeout mode is not specified.
;					$LOW_STRIKEOUT_BOLD(4); specifies to strike out the characters with a bold line.
;					$LOW_STRIKEOUT_SLASH(5); specifies to strike out the characters with slashes.
;					$LOW_STRIKEOUT_X(6); specifies to strike out the characters with X's.
; Related .......:_LOWriter_DirFrmtClear, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtStrikeOut(ByRef $oSelection, $bWordOnly = Null, $bStrikeOut = Null, $iStrikeLineStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_AnyAreDefault($bWordOnly, $bStrikeOut, $iStrikeLineStyle) Then
		If ($bWordOnly = Default) Then
			$oSelection.setPropertyToDefault("CharWordMode")
			$bWordOnly = Null
		EndIf

		If ($bStrikeOut = Default) Then
			$oSelection.setPropertyToDefault("CharCrossedOut")
			$bStrikeOut = Null
		EndIf

		If ($iStrikeLineStyle = Default) Then
			$oSelection.setPropertyToDefault("CharStrikeout")
			$iStrikeLineStyle = Null
		EndIf

		If __LOWriter_VarsAreNull($bWordOnly, $bStrikeOut, $iStrikeLineStyle) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharStrikeOut($oSelection, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DirFrmtStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DirFrmtUnderLine
; Description ...: Set and retrieve the UnderLine settings by Direct Formatting.
; Syntax ........: _LOWriter_DirFrmtUnderLine(Byref $oSelection[, $bWordOnly = Null[, $iUnderLineStyle = Null[, $bULHasColor = Null[, $iULColor = Null]]]])
; Parameters ....: $oSelection             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						or retrieval function, Or A Paragraph Object, or other Object containing a selection
;				   +						of text.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not underlined.
;                  $iUnderLineStyle     - [optional] an integer value. Default is Null. The style of the Underline line, see
;				   +									constants listed below.
;                  $bULHasColor         - [optional] a boolean value. Default is Null. Whether the underline is colored, must
;				   +						be set to true in order to set the underline color.
;                  $iULColor            - [optional] an integer value. Default is Null. The color of the underline, set in Long
;				   +						integer format. Can be one of the constants below or a custom value. Set to
;				   +						$LOW_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSelection not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSelection does not support any of the following:
;				   +								"com.sun.star.text.Paragraph";"TextPortion"; "TextCursor"; "TextViewCursor".
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iUnderLineStyle not an Integer, or less than 0 or greater than 18. Check
;				   +									the Constants list.
;				   @Error 1 @Extended 6 Return 0 = $bULHasColor not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iULColor not an Integer, or less than -1 or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $iUnderLineStyle
;				   |								4 = Error setting $ULHasColor
;				   |								8 = Error setting $iULColor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended 0 Return 2 = Success. One or more parameter was set to Default, and rest of parameters
;				   +								were set to Null. Direct formatting has been successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Direct formatting is, just as the name indicates, directly applying settings to a selection of text, it is
;						messy to deal with both by proxy (such as by Autoit automation) and directly in the document, and is
;						generally not recommended to use. Use at your own risk. Character and Paragraph styles are recommended
;						instead.
; 				   Retrieving current settings in any Direct formatting functions may be inaccurate as multiple different
;						settings could be selected at once, which would result in a return of 0, false, null, etc.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Call a Parameter with Default keyword to clear direct formatting for that setting. Note: Underline style,
;						Color and $bHasColor all reset together.
;					Note: $bULHasColor must be set to true in order to set the underline color.
; UnderLine line style Constants: $LOW_UNDERLINE_NONE(0),
;									$LOW_UNDERLINE_SINGLE(1),
;									$LOW_UNDERLINE_DOUBLE(2),
;									$LOW_UNDERLINE_DOTTED(3),
;									$LOW_UNDERLINE_DONT_KNOW(4),
;									$LOW_UNDERLINE_DASH(5),
;									$LOW_UNDERLINE_LONG_DASH(6),
;									$LOW_UNDERLINE_DASH_DOT(7),
;									$LOW_UNDERLINE_DASH_DOT_DOT(8),
;									$LOW_UNDERLINE_SML_WAVE(9),
;									$LOW_UNDERLINE_WAVE(10),
;									$LOW_UNDERLINE_DBL_WAVE(11),
;									$LOW_UNDERLINE_BOLD(12),
;									$LOW_UNDERLINE_BOLD_DOTTED(13),
;									$LOW_UNDERLINE_BOLD_DASH(14),
;									$LOW_UNDERLINE_BOLD_LONG_DASH(15),
;									$LOW_UNDERLINE_BOLD_DASH_DOT(16),
;									$LOW_UNDERLINE_BOLD_DASH_DOT_DOT(17),
;									$LOW_UNDERLINE_BOLD_WAVE(18)
; Color Constants: $LOW_COLOR_OFF(-1),
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
; Related .......: _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong, _LOWriter_DirFrmtClear,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DirFrmtUnderLine(ByRef $oSelection, $bWordOnly = Null, $iUnderLineStyle = Null, $bULHasColor = Null, $iULColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_DirFrmtCheck($oSelection) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_AnyAreDefault($bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor) Then
		If ($bWordOnly = Default) Then
			$oSelection.setPropertyToDefault("CharWordMode")
			$bWordOnly = Null
		EndIf

		If ($iUnderLineStyle = Default) Then
			$oSelection.setPropertyToDefault("CharUnderline")
			$iUnderLineStyle = Null
		EndIf

		If ($bULHasColor = Default) Then
			$oSelection.setPropertyToDefault("CharUnderlineHasColor")
			$bULHasColor = Null
		EndIf

		If ($iULColor = Default) Then
			$oSelection.setPropertyToDefault("CharUnderlineColor")
			$iULColor = Null
		EndIf

		If __LOWriter_VarsAreNull($bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	$vReturn = __LOWriter_CharUnderLine($oSelection, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)
	Return SetError(@error, @extended, $vReturn)

EndFunc   ;==>_LOWriter_DirFrmtUnderLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteDelete
; Description ...: Delete a Endnote.
; Syntax ........: _LOWriter_EndnoteDelete(Byref $oEndNote)
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous Endnote insert, or retrieval
;				   +							function.
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
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous Endnote insert, or retrieval
;				   +							function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oEndNote not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully returned the Endnote Anchor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Anchor cursor returned is just a Text Cursor placed at the anchor's position.
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert, _LOWriter_CursorMove, _LOWriter_DocGetString,
;					_LOWriter_DocInsertString
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
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous Endnote insert, or retrieval
;				   +							function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the Endnote.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $oCursor is a Table cursor type, not supported.
;				   @Error 1 @Extended 5 Return 0 = $oCursor currently located in a Frame, Footnote, Endnote, or Header/ Footer
;				   +									cannot insert a Endnote in those data types.
;				   @Error 1 @Extended 6 Return 0 = $oCursor located in unknown data type.
;				   @Error 1 @Extended 7 Return 0 = $sLabel not a string.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 =  Error creating "com.sun.star.text.Endnote" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully inserted a new Endnote, returning Endnote Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Endnote cannot be inserted into a Frame, a Footnote, a Endnote, or the Header/ Footer.
; Related .......: _LOWriter_EndnoteDelete,  _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor,
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
			Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0) ;Unsupported cursor type.
		Case $LOW_CURDATA_BODY_TEXT, $LOW_CURDATA_CELL
			$oEndNote = $oDoc.createInstance("com.sun.star.text.Endnote")
			If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

		Case Else
			Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0) ;Unknown Cursor type.
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
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous Endnote insert, or retrieval
;				   +							function.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the Endnote. Set
;				   +							to "" for automatic numbering.
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
;				   @Error 0 @Extended 2 Return String = Success. $sLabel set to Null, current Endnote AutoNumbering number
;				   +									returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
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
		;If Label is blank, return the AutoNumbering Number.
		If ($oEndNote.Label() = "") Then Return SetError($__LOW_STATUS_SUCCESS, 2, $oEndNote.Anchor.String())

		;Else return the Label.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $iNumFormat            - [optional] an integer value. Default is Null. The numbering format to use for
;				   +						Endnote numbering. See Constants.
;                  $iStartAt            - [optional] an integer value. Default is Null. The Number to begin Endnote counting
;				   +							from, Min. 1, Max 9999.
;                  $sBefore             - [optional] a string value. Default is Null. The text to display before a Endnote
;				   +							number in the note text.
;                  $sAfter              - [optional] a string value. Default is Null. The text to display after a Endnote
;				   +							number in the note text.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iNumFormat not an Integer, or Less than 0 or greater than 71. See Constants.
;				   @Error 1 @Extended 3 Return 0 = $iStartAt not an integer, less than 1 or greater than 9999.
;				   @Error 1 @Extended 4 Return 0 = $sBefore not a String.
;				   @Error 1 @Extended 5 Return 0 = $sAfter not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $iNumFormat
;				   |								2 = Error setting $iStartAt
;				   |								4 = Error setting $sBefore
;				   |								8 = Error setting $sAfter
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Numbering Format Constants: $LOW_NUM_STYLE_CHARS_UPPER_LETTER(0), Numbering is put in upper case letters. ("A, B, C, D)
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

	;0 Based -- Minus 1
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sParagraph          - [optional] a string value. Default is Null. The Endnote Text Paragraph Style.
;                  $sPage               - [optional] a string value. Default is Null. The Page Style to use for the Endnote
;				   +							pages.
;                  $sTextArea           - [optional] a string value. Default is Null. The Character Style to use for the Endnote
;				   +						anchor in the document text.
;                  $sEndnoteArea        - [optional] a string value. Default is Null. The Character Style to use for the Endnote
;				   +						number in the Endnote text.
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
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $sParagraph
;				   |								2 = Error setting $sPage
;				   |								4 = Error setting $sTextArea
;				   |								8 = Error setting $sEndnoteArea
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Endnotes Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully searched for Endnotes, none contained in document.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for Endnotes, Returning Array of Endnote
;				   +										Objects. @Extended set to number found.
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
; Name ..........: _LOWriter_FieldAuthorInsert
; Description ...: Insert a Author Field.
; Syntax ........: _LOWriter_FieldAuthorInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sAuthor = Null[, $bFullName = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author Name to insert. Note, $bIsFixed
;				   +									must be set to True for this value to stay the same as set.
;                  $bFullName           - [optional] a boolean value. Default is Null. If True, displays the full name. Else
;				   +									Initials. For a Fixed custom name, this does nothing.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sAuthor not a String.
;				   @Error 1 @Extended 7 Return 0 = $bFullName not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.Author" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Author field, returning Author Field
;				   +										Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldAuthorModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldAuthorInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sAuthor = Null, $bFullName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAuthField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oAuthField = $oDoc.createInstance("com.sun.star.text.TextField.Author")
	If Not IsObj($oAuthField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oAuthField.IsFixed = $bIsFixed
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oAuthField.Content = $sAuthor
	EndIf

	If ($bFullName <> Null) Then
		If Not IsBool($bFullName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oAuthField.FullName = $bFullName
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oAuthField, $bOverwrite)

	If ($sAuthor <> Null) Then ;Sometimes Author Disappears upon Insertion, make a check to re-set the Author value.
		If $oAuthField.Content <> $sAuthor And ($oAuthField.IsFixed() = True) Then $oAuthField.Content = $sAuthor
	EndIf

	If ($oAuthField.IsFixed() = False) Then $oAuthField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oAuthField)
EndFunc   ;==>_LOWriter_FieldAuthorInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldAuthorModify
; Description ...: Set or Retrieve a Author Field's settings.
; Syntax ........: _LOWriter_FieldAuthorModify(Byref $oAuthField[, $bIsFixed = Null[, $sAuthor = Null[, $bFullName = Null]]])
; Parameters ....: $oAuthField          - [in/out] an object. A Author field Object from a previous Insert or retrieval
;				   +									function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author Name to insert. Note, $bIsFixed
;				   +									must be set to True for this value to stay the same as set.
;                  $bFullName           - [optional] a boolean value. Default is Null. If True, displays the full name. Else
;				   +									Initials.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oAuthField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sAuthor not a String.
;				   @Error 1 @Extended 4 Return 0 = $bFullName not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sAuthor
;				   |								4 = Error setting $bFullName
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldAuthorInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldAuthorModify(ByRef $oAuthField, $bIsFixed = Null, $sAuthor = Null, $bFullName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avAuth[3]

	If Not IsObj($oAuthField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sAuthor, $bFullName) Then
		__LOWriter_ArrayFill($avAuth, $oAuthField.IsFIxed(), $oAuthField.Content(), $oAuthField.FullName())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avAuth)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oAuthField.IsFIxed = $bIsFixed
		$iError = ($oAuthField.IsFIxed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oAuthField.Content = $sAuthor
		$iError = ($oAuthField.Content() = $sAuthor) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bFullName <> Null) Then
		If Not IsBool($bFullName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oAuthField.FullName = $bFullName
		$iError = ($oAuthField.FullName() = $bFullName) ? $iError : BitOR($iError, 4)
	EndIf

	$oAuthField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldAuthorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldChapterInsert
; Description ...: Insert a Chapter Field.
; Syntax ........: _LOWriter_FieldChapterInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $iChapFrmt = Null[, $iLevel = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iChapFrmt           - [optional] an integer value. Default is Null. The Display format for the Chapter Field.
;				   +									See Constants.
;                  $iLevel              - [optional] an integer value. Default is Null. The Chapter level to display. Min. 1,
;				   +									Max 10.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iChapFrmt not an integer, less than 0 or greater than 4. See Constants.
;				   @Error 1 @Extended 6 Return 0 = $iLevel not an Integer, less than 1 or greater than 10.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.Chapter" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Chapter field, returning Chapter Field
;				   +										Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Chapter Format Constants: $LOW_FIELD_CHAP_FRMT_NAME(0), The title of the chapter is displayed.
;							$LOW_FIELD_CHAP_FRMT_NUMBER(1), The number including prefix and suffix of the chapter is displayed.
;							$LOW_FIELD_CHAP_FRMT_NAME_NUMBER(2), The title and number, with prefix and suffix of the chapter are
;								displayed.
;							$LOW_FIELD_CHAP_FRMT_NO_PREFIX_SUFFIX(3), The name and number of the chapter are displayed.
;							$LOW_FIELD_CHAP_FRMT_DIGIT(4), The number of the chapter is displayed.
; Related .......: _LOWriter_FieldChapterModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldChapterInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $iChapFrmt = Null, $iLevel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oChapField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oChapField = $oDoc.createInstance("com.sun.star.text.TextField.Chapter")
	If Not IsObj($oChapField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($iChapFrmt <> Null) Then
		If Not __LOWriter_IntIsBetween($iChapFrmt, $LOW_FIELD_CHAP_FRMT_NAME, $LOW_FIELD_CHAP_FRMT_DIGIT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oChapField.ChapterFormat = $iChapFrmt
	EndIf

	If ($iLevel <> Null) Then
		If Not __LOWriter_IntIsBetween($iLevel, 1, 10) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oChapField.Level = ($iLevel - 1) ;Level is 0 Based
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oChapField, $bOverwrite)

	$oChapField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oChapField)
EndFunc   ;==>_LOWriter_FieldChapterInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldChapterModify
; Description ...: Set or Retrieve a Chapter Field's settings.
; Syntax ........: _LOWriter_FieldChapterModify(Byref $oChapField[, $iChapFrmt = Null[, $iLevel = Null]])
; Parameters ....: $oChapField          - [in/out] an object. A Chapter field Object from a previous Insert or retrieval
;				   +									function.
;                  $iChapFrmt           - [optional] an integer value. Default is Null. The Display format for the Chapter Field.
;				   +									See Constants.
;                  $iLevel              - [optional] an integer value. Default is Null. The Chapter level to display. Min. 1,
;				   +									Max 10.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oChapField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iChapFrmt not an integer, less than 0 or greater than 4. See Constants.
;				   @Error 1 @Extended 3 Return 0 = $iLevel not an Integer, less than 1o r greater than 10.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $iChapFrmt
;				   |								2 = Error setting $iLevel
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;Chapter Format Constants: $LOW_FIELD_CHAP_FRMT_NAME(0), The title of the chapter is displayed.
;							$LOW_FIELD_CHAP_FRMT_NUMBER(1), The number including prefix and suffix of the chapter is displayed.
;							$LOW_FIELD_CHAP_FRMT_NAME_NUMBER(2), The title and number, with prefix and suffix of the chapter are
;								displayed.
;							$LOW_FIELD_CHAP_FRMT_NO_PREFIX_SUFFIX(3), The name and number of the chapter are displayed.
;							$LOW_FIELD_CHAP_FRMT_DIGIT(4), The number of the chapter is displayed.
; Related .......: _LOWriter_FieldChapterInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldChapterModify(ByRef $oChapField, $iChapFrmt = Null, $iLevel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiChap[2]

	If Not IsObj($oChapField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iChapFrmt, $iLevel) Then
		__LOWriter_ArrayFill($aiChap, $oChapField.ChapterFormat(), ($oChapField.Level() + 1)) ;Level is 0 Based -- Add 1 to make it like L.O. UI
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiChap)
	EndIf

	If ($iChapFrmt <> Null) Then
		If Not __LOWriter_IntIsBetween($iChapFrmt, $LOW_FIELD_CHAP_FRMT_NAME, $LOW_FIELD_CHAP_FRMT_DIGIT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oChapField.ChapterFormat = $iChapFrmt
		$iError = ($oChapField.ChapterFormat() = $iChapFrmt) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iLevel <> Null) Then
		If Not __LOWriter_IntIsBetween($iLevel, 1, 10) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oChapField.Level = ($iLevel - 1) ;Level is 0 Based
		$iError = ($oChapField.Level() = ($iLevel - 1)) ? $iError : BitOR($iError, 2)
	EndIf

	$oChapField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldChapterModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCombCharInsert
; Description ...: Insert a Combined Character Field.
; Syntax ........: _LOWriter_FieldCombCharInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $sCharacters = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $sCharacters         - [optional] a string value. Default is Null. The Characters to insert in a combined
;				   +									character field. Max String Length = 6.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sCharacters not a String.
;				   @Error 1 @Extended 6 Return 0 = $sCharacters longer than 6 characters.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.CombinedCharacters" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Combined Character field, returning
;				   +										Combined Character Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldCombCharModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCombCharInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sCharacters = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCombCharField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oCombCharField = $oDoc.createInstance("com.sun.star.text.TextField.CombinedCharacters")
	If Not IsObj($oCombCharField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($sCharacters <> Null) Then
		If Not IsString($sCharacters) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		If (StringLen($sCharacters) > 6) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oCombCharField.Content = $sCharacters
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oCombCharField, $bOverwrite)

	$oCombCharField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCombCharField)
EndFunc   ;==>_LOWriter_FieldCombCharInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCombCharModify
; Description ...: Set or Retrieve a Combined Character Field's settings.
; Syntax ........: _LOWriter_FieldCombCharModify(Byref $oCombCharField[, $sCharacters = Null])
; Parameters ....: $oCombCharField      - [in/out] an object. A Combined Character field Object from a previous Insert or
;				   +									retrieval function.
;                  $sCharacters         - [optional] a string value. Default is Null. The Characters to insert in a combined
;				   +									character field. Max String Length = 6.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCharacters not a String.
;				   @Error 1 @Extended 3 Return 0 = $sCharacters longer than 6 characters.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1
;				   |								1 = Error setting $sCharacters
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning current
;				   +								Combined Characters value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldCombCharInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCombCharModify(ByRef $oCombCharField, $sCharacters = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oCombCharField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sCharacters) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oCombCharField.Content())

	If Not IsString($sCharacters) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (StringLen($sCharacters) > 6) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oCombCharField.Content = $sCharacters
	$iError = ($oCombCharField.Content() = $sCharacters) ? $iError : BitOR($iError, 1)

	$oCombCharField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldCombCharModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCommentInsert
; Description ...: Insert a Comment field into a document at a cursor's position.
; Syntax ........: _LOWriter_FieldCommentInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $sContent = Null[, $sAuthor = Null[, $tDateStruct = Null[, $sInitials = Null[, $sName = Null[, $bResolved = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $sContent            - [optional] a string value. Default is Null. The content of the comment.
;                  $sAuthor             - [optional] a string value. Default is Null. The author of the comment.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display for the comment,
;				   +								created previously by _LOWriter_DateStructCreate. If left as Null, the
;				   +								current date is used.
;                  $sInitials           - [optional] a string value. Default is Null. The Initials of the creator.
;				   +								Libre Offive version 4.0 and up only.
;                  $sName               - [optional] a string value. Default is Null. The name of the creator.
;				   +								Libre Offive version 4.0 and up only.
;                  $bResolved           - [optional] a boolean value. Default is Null. If True, the comment is marked as
;				   +								resolved.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sContent not a String.
;				   @Error 1 @Extended 6 Return 0 = $sAuthor not a String.
;				   @Error 1 @Extended 7 Return 0 = $tDateStruct not an Object.
;				   @Error 1 @Extended 8 Return 0 = $sInitials not a String.
;				   @Error 1 @Extended 9 Return 0 = $sName not a String.
;				   @Error 1 @Extended 10 Return 0 = $bResolved not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.Annotation" Object.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office Version lower than 4.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted comment field, returning Comment
;				   +										Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldCommentModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DateStructCreate _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCommentInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sContent = Null, $sAuthor = Null, $tDateStruct = Null, $sInitials = Null, $sName = Null, $bResolved = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCommentField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oCommentField = $oDoc.createInstance("com.sun.star.text.TextField.Annotation")
	If Not IsObj($oCommentField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oCommentField.Content = $sContent
	Else
		$oCommentField.Content = " " ;If Content is Blank, Comment/Annotation will disappear.
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oCommentField.Author = $sAuthor
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oCommentField.DateTimeValue = $tDateStruct
	Else
		$oCommentField.DateTimeValue = _LOWriter_DateStructCreate()
	EndIf

	If ($sInitials <> Null) Then
		If Not IsString($sInitials) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$oCommentField.Initials = $sInitials
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$oCommentField.Name = $sName
	EndIf

	If ($bResolved <> Null) Then
		If Not IsBool($bResolved) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
		$oCommentField.Resolved = $bResolved
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oCommentField, $bOverwrite)

	$oCommentField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCommentField)
EndFunc   ;==>_LOWriter_FieldCommentInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCommentModify
; Description ...: Set or retrieve Comment settings.
; Syntax ........: _LOWriter_FieldCommentModify(Byref $oDoc, Byref $oCommentField[, $sContent = Null[, $sAuthor = Null[, $tDateStruct = Null[, $sInitials = Null[, $sName = Null[, $bResolved = Null]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCommentField       - [in/out] an object. A Comment field Object from a previous Insert or retrieval
;				   +									function.
;                  $sContent            - [optional] a string value. Default is Null. The content of the comment.
;                  $sAuthor             - [optional] a string value. Default is Null. The author of the comment.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display for the comment,
;				   +								created previously by _LOWriter_DateStructCreate.
;                  $sInitials           - [optional] a string value. Default is Null. The Initials of the creator.
;				   +								Libre Offive version 4.0 and up only.
;                  $sName               - [optional] a string value. Default is Null. The name of the creator.
;				   +								Libre Offive version 4.0 and up only.
;                  $bResolved           - [optional] a boolean value. Default is Null. If True, the comment is marked as
;				   +								resolved.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCommentField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sContent not a String.
;				   @Error 1 @Extended 4 Return 0 = $sAuthor not a String.
;				   @Error 1 @Extended 5 Return 0 = $tDateStruct not an Object.
;				   @Error 1 @Extended 6 Return 0 = $sInitials not a String.
;				   @Error 1 @Extended 7 Return 0 = $sName not a String.
;				   @Error 1 @Extended 8 Return 0 = $bResolved not a Boolean.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office Version lower than 4.0.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16, 32
;				   |								1 = Error setting $sContent
;				   |								2 = Error setting $sAuthor
;				   |								4 = Error setting $tDateStruct
;				   |								8 = Error setting $sInitials
;				   |								16 = Error setting $sName
;				   |								32 = Error setting $bResolved
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array If L.O. version is less than 4.0, else a 6
;				   +								Element Array with values in order of function
;				   +								parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldCommentInsert, _LOWriter_FieldsGetList, _LOWriter_DateStructCreate _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCommentModify(ByRef $oDoc, ByRef $oCommentField, $sContent = Null, $sAuthor = Null, $tDateStruct = Null, $sInitials = Null, $sName = Null, $bResolved = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avAnnot[4]
	Local $bRefresh = False

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCommentField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sContent, $sAuthor, $tDateStruct, $sInitials, $sName, $bResolved) Then
		If __LOWriter_VersionCheck(4.0) Then
			__LOWriter_ArrayFill($avAnnot, $oCommentField.Content(), $oCommentField.Author(), $oCommentField.DateTimeValue(), $oCommentField.Initials(), _
					$oCommentField.Name(), $oCommentField.Resolved())
		Else
			__LOWriter_ArrayFill($avAnnot, $oCommentField.Content(), $oCommentField.Author(), $oCommentField.DateTimeValue(), $oCommentField.Resolved())
		EndIf
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avAnnot)
	EndIf

	If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oCommentField.Content = $sContent
		$iError = ($oCommentField.Content() = $sContent) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oCommentField.Author = $sAuthor
		$iError = ($oCommentField.Author() = $sAuthor) ? $iError : BitOR($iError, 2)
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oCommentField.DateTimeValue = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oCommentField.DateTimeValue(), $tDateStruct)) ? $iError : BitOR($iError, 4)
		$bRefresh = True
	EndIf

	If ($sInitials <> Null) Then
		If Not IsString($sInitials) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$oCommentField.Initials = $sInitials
		$iError = ($oCommentField.Initials() = $sInitials) ? $iError : BitOR($iError, 8)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$oCommentField.Name = $sName
		$iError = ($oCommentField.Name = $sName) ? $iError : BitOR($iError, 16)
	EndIf

	If ($bResolved <> Null) Then
		If Not IsBool($bResolved) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oCommentField.Resolved = $bResolved
		$iError = ($oCommentField.Resolved() = $bResolved) ? $iError : BitOR($iError, 32)
		$bRefresh = True
	EndIf

	If ($bRefresh = True) Then
;~ $oCommentField = $oDoc.createInstance("com.sun.star.text.TextField.Annotation")
;~ If Not IsObj($oCommentField) Then Return  SetError($__LOW_STATUS_INIT_ERROR,1,0)
		$oDoc.Text.createTextCursorByRange($oCommentField.Anchor()).Text.insertTextContent($oCommentField.Anchor(), $oCommentField, True)
	EndIf



	$oCommentField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldCommentModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCondTextInsert
; Description ...: Insert a Conditional Text Field.
; Syntax ........: _LOWriter_FieldCondTextInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $sCondition = Null[, $sThen = Null[, $sElse = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $sCondition          - [optional] a string value. Default is Null. The condition to test.
;                  $sThen               - [optional] a string value. Default is Null. The text to display if the condition is
;				   +									true.
;                  $sElse               - [optional] a string value. Default is Null. The text to display if the condition is
;				   +									False.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sCondition not a String.
;				   @Error 1 @Extended 6 Return 0 = $sThen not a String.
;				   @Error 1 @Extended 7 Return 0 = $sElse not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.ConditionalText" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Conditional Text field, returning
;				   +										Conditional Text Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldCondTextModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCondTextInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sCondition = Null, $sThen = Null, $sElse = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCondTextField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oCondTextField = $oDoc.createInstance("com.sun.star.text.TextField.ConditionalText")
	If Not IsObj($oCondTextField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oCondTextField.Condition = $sCondition
	EndIf

	If ($sThen <> Null) Then
		If Not IsString($sThen) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oCondTextField.TrueContent = $sThen
	EndIf

	If ($sElse <> Null) Then
		If Not IsString($sElse) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oCondTextField.FalseContent = $sElse
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oCondTextField, $bOverwrite)

	$oCondTextField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCondTextField)
EndFunc   ;==>_LOWriter_FieldCondTextInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCondTextModify
; Description ...: Set or Retrieve a Conditional Text Field's settings.
; Syntax ........: _LOWriter_FieldCondTextModify(Byref $oCondTextField[, $sCondition = Null[, $sThen = Null[, $sElse = Null]]])
; Parameters ....: $oCondTextField      - [in/out] an object. A Conditional Text field Object from a previous Insert or retrieval
;				   +									function.
;                  $sCondition          - [optional] a string value. Default is Null. The condition to test.
;                  $sThen               - [optional] a string value. Default is Null. The text to display if the condition is
;				   +									 true.
;                  $sElse               - [optional] a string value. Default is Null. The text to display if the condition is
;				   +									False.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCondition not a String.
;				   @Error 1 @Extended 3 Return 0 = $sThen not a String.
;				   @Error 1 @Extended 4 Return 0 = $sElse not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $sCondition
;				   |								2 = Error setting $sThen
;				   |								4 = Error setting $sElse
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters,
;				   +								with an additional parameter in the last element to indicate if the
;				   +								condition is evaluated as True or not.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldCondTextInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCondTextModify(ByRef $oCondTextField, $sCondition = Null, $sThen = Null, $sElse = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avCond[4]

	If Not IsObj($oCondTextField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sCondition, $sThen, $sElse) Then
		__LOWriter_ArrayFill($avCond, $oCondTextField.Condition(), $oCondTextField.TrueContent(), $oCondTextField.FalseContent(), _
				($oCondTextField.IsConditionTrue()) ? False : True) ; IsConditionTrue is Backwards.
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avCond)
	EndIf

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oCondTextField.Condition = $sCondition
		$iError = ($oCondTextField.Condition() = $sCondition) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sThen <> Null) Then
		If Not IsString($sThen) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oCondTextField.TrueContent = $sThen
		$iError = ($oCondTextField.TrueContent() = $sThen) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sElse <> Null) Then
		If Not IsString($sElse) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oCondTextField.FalseContent = $sElse
		$iError = ($oCondTextField.FalseContent() = $sElse) ? $iError : BitOR($iError, 4)
	EndIf

	$oCondTextField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldCondTextModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCurrentDisplayGet
; Description ...: Retrieve the current data displayed by a field.
; Syntax ........: _LOWriter_FieldCurrentDisplayGet(Byref $oField)
; Parameters ....: $oField              - [in/out] an object. A Field Object returned from a previous insert or retrieval
;				   +							function.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oField not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returning current Field display content in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note, a Comment Field will return an empty string, use the Comment Field function to retrieve the current
;					comment content. A DocInfoComments field will work with this function however.
;					Note: This will work for most Fields, but not all. Check and see which will work and which wont.
; Related .......: _LOWriter_FieldsGetList, _LOWriter_FieldsAdvGetList, _LOWriter_FieldsDocInfoGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCurrentDisplayGet(ByRef $oField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sPresentation

	If Not IsObj($oField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($oField.supportsService("com.sun.star.text.textfield.ConditionalText")) Then ;COnditional Text Fields don't update "CurrentPresentation" setting,
		;so acquire the current display based on whether the condition is true or not.
		$sPresentation = ($oField.IsConditionTrue() = False) ? $oField.TrueContent() : $oField.FalseContent()
	Else
		$sPresentation = $oField.CurrentPresentation()
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, $sPresentation)
EndFunc   ;==>_LOWriter_FieldCurrentDisplayGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDateTimeInsert
; Description ...: Insert a Date and/or Time Field.
; Syntax ........: _LOWriter_FieldDateTimeInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $tDateStruct = Null[, $bIsDate = Null[, $iOffset = Null[, $iDateFormatKey = Null]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display for the comment,
;				   +								created previously by _LOWriter_DateStructCreate.
;                  $bIsDate             - [optional] a boolean value. Default is Null. If set to True, field is considered as
;				   +									containing a Date, $iOffset will be evaluated in Days. Else False, Field
;				   +									will be considered as containing a Time, $iOffset will be evaluated in
;				   +									minutes.
;                  $iOffset             - [optional] an integer value. Default is Null. The offset to apply to the date, either
;				   +									in Minutes or Days, depending on the current $bIsDate setting.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from
;				   +									a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList
;				   +									function.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $tDateStruct not an Object.
;				   @Error 1 @Extended 7 Return 0 = $bIsDate not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $iOffset not an Integer.
;				   @Error 1 @Extended 9 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 10 Return 0 = $iDateFormatKey not found in current Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.DateTime" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Date/Time field, returning
;				   +										Date/Time Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDateTimeModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DateStructCreate, _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDateTimeInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $tDateStruct = Null, $bIsDate = Null, $iOffset = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDateTimeField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDateTimeField = $oDoc.createInstance("com.sun.star.text.TextField.DateTime")
	If Not IsObj($oDateTimeField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDateTimeField.IsFixed = $bIsFixed
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDateTimeField.DateTimeValue = $tDateStruct
	EndIf

	If ($bIsDate <> Null) Then
		If Not IsBool($bIsDate) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oDateTimeField.IsDate = $bIsDate
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oDateTimeField.Adjust = ($oDateTimeField.IsDate() = True) ? Int((1440 * $iOffset)) : $iOffset
		;If IsDate = True, Then Calculate number of minutes in a day (1440) times number of days to off set the Date/ Value,
		;else, just set it to Number of minutes called.
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
		$oDateTimeField.NumberFormat = $iDateFormatKey
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDateTimeField, $bOverwrite)

	If ($tDateStruct <> Null) Then ;Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If (__LOWriter_DateStructCompare($oDateTimeField.DateTimeValue(), $tDateStruct) = False) And ($oDateTimeField.IsFixed() = True) Then $oDateTimeField.DateTimeValue = $tDateStruct
	EndIf

	If ($oDateTimeField.IsFixed() = False) Then $oDateTimeField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDateTimeField)
EndFunc   ;==>_LOWriter_FieldDateTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDateTimeModify
; Description ...: Set or Retrieve a Date/Time Field's settings.
; Syntax ........: _LOWriter_FieldDateTimeModify(Byref $oDoc, Byref $oDateTimeField[, $bIsFixed = Null[, $tDateStruct = Null[, $bIsDate = Null[, $iOffset = Null[, $iDateFormatKey = Null]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oDateTimeField      - [in/out] an object. A Date/Time field Object from a previous Insert or retrieval
;				   +									function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display for the comment,
;				   +								created previously by _LOWriter_DateStructCreate.
;                  $bIsDate             - [optional] a boolean value. Default is Null. If set to True, field is considered as
;				   +									containing a Date, $iOffset will be evaluated in Days. Else False, Field
;				   +									will be considered as containing a Time, $iOffset will be evaluated in
;				   +									minutes.
;                  $iOffset             - [optional] an integer value. Default is Null. The offset to apply to the date, either
;				   +									in Minutes or Days, depending on the current $bIsDate setting.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from
;				   +									a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList
;				   +									function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $tDateStruct not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bIsDate not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iOffset not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iDateFormatKey not found in current Document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $tDateStruct
;				   |								4 = Error setting $bIsDate
;				   |								8 = Error setting $iOffset
;				   |								16 = Error setting $iDateFormatKey
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDateTimeInsert, _LOWriter_FieldsGetList, _LOWriter_DateStructCreate,
;					_LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDateTimeModify(ByRef $oDoc, ByRef $oDateTimeField, $bIsFixed = Null, $tDateStruct = Null, $bIsDate = Null, $iOffset = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avDateTime[5]

	If Not IsObj($oDateTimeField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $tDateStruct, $bIsDate, $iOffset, $iDateFormatKey) Then
		;Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		;fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		;from the value.
		$iNumberFormat = $oDateTimeField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : $iNumberFormat

		__LOWriter_ArrayFill($avDateTime, $oDateTimeField.IsFixed(), $oDateTimeField.DateTimeValue(), $oDateTimeField.IsDate(), _
				($oDateTimeField.IsDate() = True) ? Int(($oDateTimeField.Adjust() / 1440)) : $oDateTimeField.Adjust(), $iNumberFormat)
		;If IsDate = True, Then Calculate number of minutes in a day (1440) divided by number of days of off set. Otherwise
		;return Number of minutes.
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDateTime)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDateTimeField.IsFixed = $bIsFixed
		$iError = ($oDateTimeField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDateTimeField.DateTimeValue = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oDateTimeField.DateTimeValue(), $tDateStruct)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bIsDate <> Null) Then
		If Not IsBool($bIsDate) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oDateTimeField.IsDate = $bIsDate
		$iError = ($oDateTimeField.IsDate() = $bIsDate) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$iOffset = ($oDateTimeField.IsDate() = True) ? Int((1440 * $iOffset)) : $iOffset
		;If IsDate = True, Then Calculate number of minutes in a day (1440) times number of days to off set the Date/ Value,
		;else, just set it to Number of minutes called.

		$oDateTimeField.Adjust = $iOffset
		$iError = ($oDateTimeField.Adjust() = $iOffset) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oDateTimeField.NumberFormat = $iDateFormatKey
		$iError = ($oDateTimeField.NumberFormat() = ($iDateFormatKey)) ? $iError : BitOR($iError, 16)
	EndIf

	If ($oDateTimeField.IsFixed() = False) Then $oDateTimeField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDateTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDelete
; Description ...: Delete a Field from a Document.
; Syntax ........: _LOWriter_FieldDelete(Byref $oDoc, Byref $oField[, $bDeleteMaster = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oField              - [in/out] an object. A Field Object from a previous Insert or retrieval function.
;                  $bDeleteMaster       - [optional] a boolean value. Default is False. If True, and the field has a Master
;				   +						Field, the MasterField (With any other dependent fields) will be deleted.
; Return values .: Success: 1.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bDeleteMaster not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving TextFieldMaster Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Field Master Array of dependent fields.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted field and the Text Master Field.
;				   @Error 0 @Extended 1 Return 1 = Success. Successfully deleted field.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:  _LOWriter_FieldsGetList, _LOWriter_FieldsAdvGetList, _LOWriter_FieldsDocInfoGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDelete(ByRef $oDoc, ByRef $oField, $bDeleteMaster = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFieldMaster
	Local $aoDependents[0]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bDeleteMaster) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If ($bDeleteMaster = True) And ($oField.TextFieldMaster.Name() <> "") Then
		$oFieldMaster = $oField.TextFieldMaster()
		If Not IsObj($oFieldMaster) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

		$aoDependents = $oFieldMaster.DependentTextFields()
		If Not IsArray($aoDependents) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

		If (UBound($aoDependents) > 0) Then
			For $i = 0 To UBound($aoDependents) - 1
				$aoDependents[$i].Anchor.Text.removeTextContent($aoDependents[$i])
				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
			Next
		EndIf

		$oFieldMaster.dispose()
		Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
	EndIf

	$oField.Anchor.Text.removeTextContent($oField)

	Return SetError($__LOW_STATUS_SUCCESS, 1, 1)
EndFunc   ;==>_LOWriter_FieldDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCommentsInsert
; Description ...: Insert a Document Information Comments Field.
; Syntax ........: _LOWriter_FieldDocInfoCommentsInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sComments = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sComments           - [optional] a string value. Default is Null. The Comments text to display, note,
;				   +									$bIsFixed must be True for this to be displayed.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sComments not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.Description" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Comments Field.
;				   +											Returning the Document Info Comments Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoCommentsModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCommentsInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sComments = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoCommentField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoCommentField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.Description")
	If Not IsObj($oDocInfoCommentField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoCommentField.IsFixed = $bIsFixed
	EndIf

	If ($sComments <> Null) Then
		If Not IsString($sComments) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoCommentField.Content = $sComments
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoCommentField, $bOverwrite)

	If ($sComments <> Null) Then ;Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oDocInfoCommentField.Content <> $sComments And ($oDocInfoCommentField.IsFixed() = True) Then $oDocInfoCommentField.Content = $sComments
	EndIf

	If ($oDocInfoCommentField.IsFixed() = False) Then $oDocInfoCommentField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoCommentField)
EndFunc   ;==>_LOWriter_FieldDocInfoCommentsInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCommentsModify
; Description ...: Set or Retrieve a Document Information Comments Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoCommentsModify(Byref $oDocInfoCommentField[, $bIsFixed = Null[, $sComments = Null]])
; Parameters ....: $oDocInfoCommentField- [in/out] an object. A Doc Info Comments field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sComments           - [optional] a string value. Default is Null. The Comments text to display, note,
;				   +									$bIsFixed must be True for this to be displayed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoCommentField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sComments not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sComments
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoCommentsInsert, _LOWriter_FieldsGetList, _LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCommentsModify(ByRef $oDocInfoCommentField, $bIsFixed = Null, $sComments = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoCom[2]

	If Not IsObj($oDocInfoCommentField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sComments) Then
		__LOWriter_ArrayFill($avDocInfoCom, $oDocInfoCommentField.IsFixed(), $oDocInfoCommentField.Content())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoCom)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoCommentField.IsFixed = $bIsFixed
		$iError = ($oDocInfoCommentField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sComments <> Null) Then
		If Not IsString($sComments) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoCommentField.Content = $sComments
		$iError = ($oDocInfoCommentField.Content() = $sComments) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoCommentField.IsFixed() = False) Then $oDocInfoCommentField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoCommentsModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCreateAuthInsert
; Description ...: Insert a Document Information Create Author Field.
; Syntax ........: _LOWriter_FieldDocInfoCreateAuthInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sAuthor = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed but be
;				   +									set to True in order for this to remain as set.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sAuthor not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.CreateAuthor" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Created By Author Field.
;				   +											Returning the Document Info Created By Author Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoCreateAuthModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocGenPropCreation
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCreateAuthInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoCreateAuthField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoCreateAuthField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.CreateAuthor")
	If Not IsObj($oDocInfoCreateAuthField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoCreateAuthField.IsFixed = $bIsFixed
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoCreateAuthField.Author = $sAuthor
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoCreateAuthField, $bOverwrite)

	If ($sAuthor <> Null) Then ;Sometimes Author Disappears upon Insertion, make a check to re-set the Author value.
		If $oDocInfoCreateAuthField.Author <> $sAuthor And ($oDocInfoCreateAuthField.IsFixed() = True) Then $oDocInfoCreateAuthField.Author = $sAuthor
	EndIf

	If ($oDocInfoCreateAuthField.IsFixed() = False) Then $oDocInfoCreateAuthField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoCreateAuthField)
EndFunc   ;==>_LOWriter_FieldDocInfoCreateAuthInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCreateAuthModify
; Description ...: Set or Retrieve a Document Information Create Author Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoCreateAuthModify(Byref $oDocInfoCreateAuthField[, $bIsFixed = Null[, $sAuthor = Null]])
; Parameters ....: $oDocInfoCreateAuthField- [in/out] an object. A Created By Author field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed but be
;				   +									set to True in order for this to remain as set.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoCreateAuthField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sAuthor not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sAuthor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoCreateAuthInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocGenPropCreation
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCreateAuthModify(ByRef $oDocInfoCreateAuthField, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoModAuth[2]

	If Not IsObj($oDocInfoCreateAuthField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sAuthor) Then
		__LOWriter_ArrayFill($avDocInfoModAuth, $oDocInfoCreateAuthField.IsFixed(), $oDocInfoCreateAuthField.Author())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoModAuth)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoCreateAuthField.IsFixed = $bIsFixed
		$iError = ($oDocInfoCreateAuthField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoCreateAuthField.Author = $sAuthor
		$iError = ($oDocInfoCreateAuthField.Author() = $sAuthor) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoCreateAuthField.IsFixed() = False) Then $oDocInfoCreateAuthField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoCreateAuthModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCreateDateTimeInsert
; Description ...: Insert a Document Information Create Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoCreateDateTimeInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iDateFormatKey = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from
;				   +									a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList
;				   +									function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iDateFormatKey not found in document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.CreateDateTime" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Created Date/Time Field.
;				   +											Returning the Document Info Created Date/Time Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoCreateDateTimeModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenPropCreation
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCreateDateTimeInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoCreateDtTmField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoCreateDtTmField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.CreateDateTime")
	If Not IsObj($oDocInfoCreateDtTmField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoCreateDtTmField.IsFixed = $bIsFixed
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oDocInfoCreateDtTmField.NumberFormat = $iDateFormatKey
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoCreateDtTmField, $bOverwrite)

	If ($oDocInfoCreateDtTmField.IsFixed() = False) Then $oDocInfoCreateDtTmField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoCreateDtTmField)
EndFunc   ;==>_LOWriter_FieldDocInfoCreateDateTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCreateDateTimeModify
; Description ...: Set or Retrieve a Document Information Create Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoCreateDateTimeModify(Byref $oDoc, Byref $oDocInfoCreateDtTmField[, $bIsFixed = Null[, $iDateFormatKey = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oDocInfoCreateDtTmField- [in/out] an object. A Created at Date/Time field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from
;				   +									a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList
;				   +									function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oDocInfoCreateDtTmField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iDateFormatKey not found in document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iDateFormatKey
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoCreateDateTimeInsert, _LOWriter_FieldsDocInfoGetList,
;					_LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenPropCreation
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCreateDateTimeModify(ByRef $oDoc, ByRef $oDocInfoCreateDtTmField, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avDocInfoCrtDate[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDocInfoCreateDtTmField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $iDateFormatKey) Then
		;Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		;fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000 from
		;the value.
		$iNumberFormat = $oDocInfoCreateDtTmField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : $iNumberFormat

		__LOWriter_ArrayFill($avDocInfoCrtDate, $oDocInfoCreateDtTmField.IsFixed(), $iNumberFormat)
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoCrtDate)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoCreateDtTmField.IsFixed = $bIsFixed
		$iError = ($oDocInfoCreateDtTmField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoCreateDtTmField.NumberFormat = $iDateFormatKey
		$iError = ($oDocInfoCreateDtTmField.NumberFormat() = $iDateFormatKey) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoCreateDtTmField.IsFixed() = False) Then $oDocInfoCreateDtTmField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoCreateDateTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoEditTimeInsert
; Description ...: Insert a Document Information Total Editing Time Field.
; Syntax ........: _LOWriter_FieldDocInfoEditTimeInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iTimeFormatKey = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $iTimeFormatKey      - [optional] an integer value. Default is Null. A Time Format Key returned from
;				   +									a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList
;				   +									function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iTimeFormatKey not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iTimeFormatKey not found in document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.EditTime" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Total Editing Time Field.
;				   +											Returning the Document Info Total Editing Time Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: _LOWriter_FieldDocInfoEditTimeModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenProp
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoEditTimeInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iTimeFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoEditTimeField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoEditTimeField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.EditTime")
	If Not IsObj($oDocInfoEditTimeField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoEditTimeField.IsFixed = $bIsFixed
	EndIf

	If ($iTimeFormatKey <> Null) Then
		If Not IsInt($iTimeFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iTimeFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oDocInfoEditTimeField.NumberFormat = $iTimeFormatKey
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoEditTimeField, $bOverwrite)

	If ($oDocInfoEditTimeField.IsFixed() = False) Then $oDocInfoEditTimeField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoEditTimeField)
EndFunc   ;==>_LOWriter_FieldDocInfoEditTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoEditTimeModify
; Description ...: Set or Retrieve a Document Information Total Editing Time Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoEditTimeModify(Byref $oDoc, Byref $oDocInfoEditTimeField[, $bIsFixed = Null[, $iTimeFormatKey = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oDocInfoEditTimeField- [in/out] an object. A Doc Info Total Editing Time field Object from a previous Insert
;				   +									or retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $iTimeFormatKey      - [optional] an integer value. Default is Null. A Time Format Key returned from
;				   +									a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList
;				   +									function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oDocInfoEditTimeField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iTimeFormatKey not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iTimeFormatKey not found in document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iTimeFormatKey
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoEditTimeInsert, _LOWriter_FieldsDocInfoGetList,
;					_LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenProp
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoEditTimeModify(ByRef $oDoc, ByRef $oDocInfoEditTimeField, $bIsFixed = Null, $iTimeFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avDocInfoEditTm[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDocInfoEditTimeField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $iTimeFormatKey) Then
		;Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		;fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		;from the value.
		$iNumberFormat = $oDocInfoEditTimeField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : $iNumberFormat

		__LOWriter_ArrayFill($avDocInfoEditTm, $oDocInfoEditTimeField.IsFixed(), $iNumberFormat)
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoEditTm)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoEditTimeField.IsFixed = $bIsFixed
		$iError = ($oDocInfoEditTimeField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iTimeFormatKey <> Null) Then
		If Not IsInt($iTimeFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		If Not IsInt($iTimeFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoEditTimeField.NumberFormat = $iTimeFormatKey
		$iError = ($oDocInfoEditTimeField.NumberFormat() = $iTimeFormatKey) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoEditTimeField.IsFixed() = False) Then $oDocInfoEditTimeField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoEditTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoKeywordsInsert
; Description ...: Insert a Document Information Keywords Field.
; Syntax ........: _LOWriter_FieldDocInfoKeywordsInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sKeywords = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sKeywords           - [optional] a string value. Default is Null. The Keywords text to display, note,
;				   +									$bIsFixed must be True for this to be displayed.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sKeywords not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.Keywords" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Keywords Field.
;				   +											Returning the Document Info Keywords Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoKeywordsModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoKeywordsInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sKeywords = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoKeywordField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoKeywordField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.KeyWords")
	If Not IsObj($oDocInfoKeywordField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoKeywordField.IsFixed = $bIsFixed
	EndIf

	If ($sKeywords <> Null) Then
		If Not IsString($sKeywords) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoKeywordField.Content = $sKeywords
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoKeywordField, $bOverwrite)

	If ($sKeywords <> Null) Then ;Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oDocInfoKeywordField.Content <> $sKeywords And ($oDocInfoKeywordField.IsFixed() = True) Then $oDocInfoKeywordField.Content = $sKeywords
	EndIf

	If ($oDocInfoKeywordField.IsFixed() = False) Then $oDocInfoKeywordField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoKeywordField)
EndFunc   ;==>_LOWriter_FieldDocInfoKeywordsInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoKeywordsModify
; Description ...: Set or Retrieve a Document Information Keywords Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoKeywordsModify(Byref $oDocInfoKeywordField[, $bIsFixed = Null[, $sKeywords = Null]])
; Parameters ....: $oDocInfoKeywordField- [in/out] an object. A Doc Info Keywords field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sKeywords           - [optional] a string value. Default is Null. The Keywords text to display, note,
;				   +									$bIsFixed must be True for this to be displayed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoKeywordField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sKeywords not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sKeywords
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoKeywordsInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoKeywordsModify(ByRef $oDocInfoKeywordField, $bIsFixed = Null, $sKeywords = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoKyWrd[2]

	If Not IsObj($oDocInfoKeywordField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sKeywords) Then
		__LOWriter_ArrayFill($avDocInfoKyWrd, $oDocInfoKeywordField.IsFixed(), $oDocInfoKeywordField.Content())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoKyWrd)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoKeywordField.IsFixed = $bIsFixed
		$iError = ($oDocInfoKeywordField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sKeywords <> Null) Then
		If Not IsString($sKeywords) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoKeywordField.Content = $sKeywords
		$iError = ($oDocInfoKeywordField.Content() = $sKeywords) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoKeywordField.IsFixed() = False) Then $oDocInfoKeywordField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoKeywordsModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoModAuthInsert
; Description ...: Insert a Document Information Modification Author Field.
; Syntax ........: _LOWriter_FieldDocInfoModAuthInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sAuthor = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;				   $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed but be
;				   +									set to True in order for this to remain as set.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sAuthor not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.ChangeAuthor" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Modified By Author Field.
;				   +											Returning the Document Info Modified By Author Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoModAuthModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocGenPropModification
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoModAuthInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoModAuthField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoModAuthField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.ChangeAuthor")
	If Not IsObj($oDocInfoModAuthField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoModAuthField.IsFixed = $bIsFixed
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoModAuthField.Author = $sAuthor
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoModAuthField, $bOverwrite)

	If ($sAuthor <> Null) Then ;Sometimes Author Disappears upon Insertion, make a check to re-set the Author value.
		If $oDocInfoModAuthField.Author <> $sAuthor And ($oDocInfoModAuthField.IsFixed() = True) Then $oDocInfoModAuthField.Author = $sAuthor
	EndIf

	If ($oDocInfoModAuthField.IsFixed() = False) Then $oDocInfoModAuthField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoModAuthField)
EndFunc   ;==>_LOWriter_FieldDocInfoModAuthInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoModAuthModify
; Description ...: Set or Retrieve a Document Information Modification Author Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoModAuthModify(Byref $oDocInfoModAuthField[, $bIsFixed = Null[, $sAuthor = Null]])
; Parameters ....: $oDocInfoModAuthField- [in/out] an object. A Modified By Author field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed but be
;				   +									set to True in order for this to remain as set.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoModAuthField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sAuthor not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sAuthor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoModAuthInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocGenPropModification
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoModAuthModify(ByRef $oDocInfoModAuthField, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoModAuth[2]

	If Not IsObj($oDocInfoModAuthField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sAuthor) Then
		__LOWriter_ArrayFill($avDocInfoModAuth, $oDocInfoModAuthField.IsFixed(), $oDocInfoModAuthField.Author())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoModAuth)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoModAuthField.IsFixed = $bIsFixed
		$iError = ($oDocInfoModAuthField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoModAuthField.Author = $sAuthor
		$iError = ($oDocInfoModAuthField.Author() = $sAuthor) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoModAuthField.IsFixed() = False) Then $oDocInfoModAuthField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoModAuthModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoModDateTimeInsert
; Description ...: Insert a Document Information Modification Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoModDateTimeInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iDateFormatKey = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from
;				   +									a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList
;				   +									function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iDateFormatKey not found in document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.ChangeDateTime" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Modified Date/Time Field.
;				   +											Returning the Document Info Modified Date/Time Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoModDateTimeModify, _LOWriter_DateFormatKeyCreate,
;					_LOWriter_DateFormatKeyList, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocGenPropModification
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoModDateTimeInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoModDtTmField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoModDtTmField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.ChangeDateTime")
	If Not IsObj($oDocInfoModDtTmField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoModDtTmField.IsFixed = $bIsFixed
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oDocInfoModDtTmField.NumberFormat = $iDateFormatKey
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoModDtTmField, $bOverwrite)

	If ($oDocInfoModDtTmField.IsFixed() = False) Then $oDocInfoModDtTmField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoModDtTmField)
EndFunc   ;==>_LOWriter_FieldDocInfoModDateTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoModDateTimeModify
; Description ...: Set or Retrieve a Document Information Modification Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoModDateTimeModify(Byref $oDoc, Byref $oDocInfoModDtTmField[, $bIsFixed = Null[, $iDateFormatKey = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oDocInfoModDtTmField- [in/out] an object. A Modified at Date/Time field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from
;				   +									a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList
;				   +									function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oDocInfoPrintAuthField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iDateFormatKey not found in document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iDateFormatKey
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoModDateTimeInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DateFormatKeyCreate,
;					_LOWriter_DateFormatKeyList, _LOWriter_DocGenPropModification
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoModDateTimeModify(ByRef $oDoc, ByRef $oDocInfoModDtTmField, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avDocInfoModDate[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDocInfoModDtTmField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $iDateFormatKey) Then
		;Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		;fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		;from the value.
		$iNumberFormat = $oDocInfoModDtTmField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : $iNumberFormat

		__LOWriter_ArrayFill($avDocInfoModDate, $oDocInfoModDtTmField.IsFixed(), $iNumberFormat)
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoModDate)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoModDtTmField.IsFixed = $bIsFixed
		$iError = ($oDocInfoModDtTmField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoModDtTmField.NumberFormat = $iDateFormatKey
		$iError = ($oDocInfoModDtTmField.NumberFormat() = $iDateFormatKey) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoModDtTmField.IsFixed() = False) Then $oDocInfoModDtTmField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoModDateTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoPrintAuthInsert
; Description ...: Insert a Document Information Last Print Author Field.
; Syntax ........: _LOWriter_FieldDocInfoPrintAuthInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sAuthor = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed but be
;				   +									set to True in order for this to remain as set.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sAuthor not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.PrintAuthor" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Printed By Author Field.
;				   +											Returning the Document Info Printed By Author Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoPrintAuthModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocGenPropPrint
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoPrintAuthInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoPrintAuthField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoPrintAuthField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.PrintAuthor")
	If Not IsObj($oDocInfoPrintAuthField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoPrintAuthField.IsFixed = $bIsFixed
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoPrintAuthField.Author = $sAuthor
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoPrintAuthField, $bOverwrite)

	If ($sAuthor <> Null) Then ;Sometimes Author Disappears upon Insertion, make a check to re-set the Author value.
		If $oDocInfoPrintAuthField.Author <> $sAuthor And ($oDocInfoPrintAuthField.IsFixed() = True) Then $oDocInfoPrintAuthField.Author = $sAuthor
	EndIf

	If ($oDocInfoPrintAuthField.IsFixed() = False) Then $oDocInfoPrintAuthField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoPrintAuthField)
EndFunc   ;==>_LOWriter_FieldDocInfoPrintAuthInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoPrintAuthModify
; Description ...: Set or Retrieve a Document Information Last Print Author Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoPrintAuthModify(Byref $oDocInfoPrintAuthField[, $bIsFixed = Null[, $sAuthor = Null]])
; Parameters ....: $oDocInfoPrintAuthField- [in/out] an object. A Printed By Author field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed but be
;				   +									set to True in order for this to remain as set.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoPrintAuthField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sAuthor not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sAuthor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoPrintAuthInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocGenPropPrint
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoPrintAuthModify(ByRef $oDocInfoPrintAuthField, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoModAuth[2]

	If Not IsObj($oDocInfoPrintAuthField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sAuthor) Then
		__LOWriter_ArrayFill($avDocInfoModAuth, $oDocInfoPrintAuthField.IsFixed(), $oDocInfoPrintAuthField.Author())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoModAuth)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoPrintAuthField.IsFixed = $bIsFixed
		$iError = ($oDocInfoPrintAuthField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoPrintAuthField.Author = $sAuthor
		$iError = ($oDocInfoPrintAuthField.Author() = $sAuthor) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoPrintAuthField.IsFixed() = False) Then $oDocInfoPrintAuthField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoPrintAuthModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoPrintDateTimeInsert
; Description ...: Insert a Document Information Last Print Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoPrintDateTimeInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iDateFormatKey = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from
;				   +									a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList
;				   +									function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iDateFormatKey not found in document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.PrintDateTime" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Printed Date/Time Field.
;				   +											Returning the Document Info Printed Date/Time Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoPrintDateTimeModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenPropPrint
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoPrintDateTimeInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoPrintDtTmField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoPrintDtTmField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.PrintDateTime")
	If Not IsObj($oDocInfoPrintDtTmField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoPrintDtTmField.IsFixed = $bIsFixed
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oDocInfoPrintDtTmField.NumberFormat = $iDateFormatKey
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoPrintDtTmField, $bOverwrite)

	If ($oDocInfoPrintDtTmField.IsFixed() = False) Then $oDocInfoPrintDtTmField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoPrintDtTmField)
EndFunc   ;==>_LOWriter_FieldDocInfoPrintDateTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoPrintDateTimeModify
; Description ...: Set or Retrieve a Document Information Last Print Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoPrintDateTimeModify(Byref $oDoc, Byref $oDocInfoPrintDtTmField[, $bIsFixed = Null[, $iDateFormatKey = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oDocInfoPrintDtTmField- [in/out] an object. A Printed at Date/Time field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from
;				   +									a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList
;				   +									function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oDocInfoPrintDtTmField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iDateFormatKey not found in document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iDateFormatKey
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoPrintDateTimeInsert,  _LOWriter_FieldsDocInfoGetList, _LOWriter_DateFormatKeyCreate,
;					_LOWriter_DateFormatKeyList, _LOWriter_DocGenPropPrint
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoPrintDateTimeModify(ByRef $oDoc, ByRef $oDocInfoPrintDtTmField, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avDocInfoPrntDate[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDocInfoPrintDtTmField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $iDateFormatKey) Then
		;Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		;fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		;from the value.
		$iNumberFormat = $oDocInfoPrintDtTmField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : $iNumberFormat

		__LOWriter_ArrayFill($avDocInfoPrntDate, $oDocInfoPrintDtTmField.IsFixed(), $iNumberFormat)
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoPrntDate)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoPrintDtTmField.IsFixed = $bIsFixed
		$iError = ($oDocInfoPrintDtTmField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoPrintDtTmField.NumberFormat = $iDateFormatKey
		$iError = ($oDocInfoPrintDtTmField.NumberFormat() = $iDateFormatKey) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoPrintDtTmField.IsFixed() = False) Then $oDocInfoPrintDtTmField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoPrintDateTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoRevNumInsert
; Description ...: Insert a Document Information Revision Number Field.
; Syntax ........: _LOWriter_FieldDocInfoRevNumInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iRevNum = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $iRevNum             - [optional] a Integer value. Default is Null. The Revision Number Integer to display,
;				   +									note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iRevNum not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.Revision" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Revision Number Field.
;				   +											Returning the Document Info Revision Number Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoRevNumModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocGenProp
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoRevNumInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iRevNum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoRevNumField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoRevNumField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.Revision")
	If Not IsObj($oDocInfoRevNumField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoRevNumField.IsFixed = $bIsFixed
	EndIf

	If ($iRevNum <> Null) Then
		If Not IsInt($iRevNum) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoRevNumField.Revision = $iRevNum
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoRevNumField, $bOverwrite)

	If ($iRevNum <> Null) Then ;Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oDocInfoRevNumField.Revision <> $iRevNum And ($oDocInfoRevNumField.IsFixed() = True) Then $oDocInfoRevNumField.Revision = $iRevNum
	EndIf

	If ($oDocInfoRevNumField.IsFixed() = False) Then $oDocInfoRevNumField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoRevNumField)
EndFunc   ;==>_LOWriter_FieldDocInfoRevNumInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoRevNumModify
; Description ...: Set or Retrieve a Document Information Revision Number Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoRevNumModify(Byref $oDocInfoRevNumField[, $bIsFixed = Null[, $iRevNum = Null]])
; Parameters ....: $oDocInfoRevNumField - [in/out] an object. A Doc Info Revision Number field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $iRevNum             - [optional] a Integer value. Default is Null. The Revision Number Integer to display,
;				   +									note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoRevNumField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $iRevNum not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iRevNum
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoRevNumInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocGenProp
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoRevNumModify(ByRef $oDocInfoRevNumField, $bIsFixed = Null, $iRevNum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoRev[2]

	If Not IsObj($oDocInfoRevNumField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $iRevNum) Then
		__LOWriter_ArrayFill($avDocInfoRev, $oDocInfoRevNumField.IsFixed(), $oDocInfoRevNumField.Revision())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoRev)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoRevNumField.IsFixed = $bIsFixed
		$iError = ($oDocInfoRevNumField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iRevNum <> Null) Then
		If Not IsInt($iRevNum) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoRevNumField.Revision = $iRevNum
		$iError = ($oDocInfoRevNumField.Revision() = $iRevNum) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoRevNumField.IsFixed() = False) Then $oDocInfoRevNumField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoRevNumModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoSubjectInsert
; Description ...: Insert a Document Information Subject Field.
; Syntax ........: _LOWriter_FieldDocInfoSubjectInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sSubject = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sSubject            - [optional] a string value. Default is Null. The Subject text to display, note,
;				   +									$bIsFixed must be True for this to be displayed.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sSubject not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.Subject" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Subject Field.
;				   +											Returning the Document Info Subject Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoSubjectModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoSubjectInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sSubject = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoSubField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoSubField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.Subject")
	If Not IsObj($oDocInfoSubField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoSubField.IsFixed = $bIsFixed
	EndIf

	If ($sSubject <> Null) Then
		If Not IsString($sSubject) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoSubField.Content = $sSubject
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoSubField, $bOverwrite)

	If ($sSubject <> Null) Then ;Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oDocInfoSubField.Content <> $sSubject And ($oDocInfoSubField.IsFixed() = True) Then $oDocInfoSubField.Content = $sSubject
	EndIf

	If ($oDocInfoSubField.IsFixed() = False) Then $oDocInfoSubField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoSubField)
EndFunc   ;==>_LOWriter_FieldDocInfoSubjectInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoSubjectModify
; Description ...: Set or Retrieve a Document Information Subject Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoSubjectModify(Byref $oDocInfoSubField[, $bIsFixed = Null[, $sSubject = Null]])
; Parameters ....: $oDocInfoSubField    - [in/out] an object. A Doc Info Subject field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sSubject            - [optional] a string value. Default is Null. The Subject text to display, note,
;				   +									$bIsFixed must be True for this to be displayed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoSubField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sSubject not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sSubject
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoSubjectInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoSubjectModify(ByRef $oDocInfoSubField, $bIsFixed = Null, $sSubject = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoSub[2]

	If Not IsObj($oDocInfoSubField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sSubject) Then
		__LOWriter_ArrayFill($avDocInfoSub, $oDocInfoSubField.IsFixed(), $oDocInfoSubField.Content())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoSub)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoSubField.IsFixed = $bIsFixed
		$iError = ($oDocInfoSubField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sSubject <> Null) Then
		If Not IsString($sSubject) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoSubField.Content = $sSubject
		$iError = ($oDocInfoSubField.Content() = $sSubject) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoSubField.IsFixed() = False) Then $oDocInfoSubField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoSubjectModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoTitleInsert
; Description ...: Insert a Document Information Title Field.
; Syntax ........: _LOWriter_FieldDocInfoTitleInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sTitle = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sTitle              - [optional] a string value. Default is Null. The Title text to display, note, $bIsFixed
;				   +									 must be True for this to be displayed.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sTitle not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.Title" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Title Field.
;				   +											Returning the Document Info Title Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoTitleModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoTitleInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sTitle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoTitleField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoTitleField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.Title")
	If Not IsObj($oDocInfoTitleField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoTitleField.IsFixed = $bIsFixed
	EndIf

	If ($sTitle <> Null) Then
		If Not IsString($sTitle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoTitleField.Content = $sTitle
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoTitleField, $bOverwrite)

	If ($sTitle <> Null) Then ;Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oDocInfoTitleField.Content <> $sTitle And ($oDocInfoTitleField.IsFixed() = True) Then $oDocInfoTitleField.Content = $sTitle
	EndIf

	If ($oDocInfoTitleField.IsFixed() = False) Then $oDocInfoTitleField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDocInfoTitleField)
EndFunc   ;==>_LOWriter_FieldDocInfoTitleInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoTitleModify
; Description ...: Set or Retrieve a Document Information Title Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoTitleModify(Byref $oDocInfoTitleField[, $bIsFixed = Null[, $sTitle = Null]])
; Parameters ....: $oDocInfoTitleField  - [in/out] an object. A Doc Info Title field Object from a previous Insert or
;				   +									retrieval function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sTitle              - [optional] a string value. Default is Null. The Title text to display, note, $bIsFixed
;				   +									 must be True for this to be displayed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoTitleField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sTitle not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sTitle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoTitleInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoTitleModify(ByRef $oDocInfoTitleField, $bIsFixed = Null, $sTitle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoTitle[2]

	If Not IsObj($oDocInfoTitleField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sTitle) Then
		__LOWriter_ArrayFill($avDocInfoTitle, $oDocInfoTitleField.IsFixed(), $oDocInfoTitleField.Content())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDocInfoTitle)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoTitleField.IsFixed = $bIsFixed
		$iError = ($oDocInfoTitleField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sTitle <> Null) Then
		If Not IsString($sTitle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoTitleField.Content = $sTitle
		$iError = ($oDocInfoTitleField.Content() = $sTitle) ? $iError : BitOR($iError, 2)
	EndIf

	If ($oDocInfoTitleField.IsFixed() = False) Then $oDocInfoTitleField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldDocInfoTitleModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFileNameInsert
; Description ...: Insert a File Name Field.
; Syntax ........: _LOWriter_FieldFileNameInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iFormat = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;				   $iFormat             - [optional] an integer value. Default is Null. The Data Format to  display. See
;				   +									Constants.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iFormat not an Integer, Less than 0, or greater than 3. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.FileName" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted File Name field, returning
;				   +										File Name Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Until at least L.O. Version 7.3.4.2, there is a bug where the wrong Path Format type is displayed when the content is set to Fixed = True.
;				   For example, $LOW_FIELD_FILENAME_NAME_AND_EXT, displays in the format of $LOW_FIELD_FILENAME_NAME.
; File Name Constants: $LOW_FIELD_FILENAME_* as defined in LibreOfficeWriter_Constants.au3
; Related .......: _LOWriter_FieldFileNameModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFileNameInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFileNameField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oFileNameField = $oDoc.createInstance("com.sun.star.text.TextField.FileName")
	If Not IsObj($oFileNameField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oFileNameField.IsFixed = $bIsFixed
	EndIf

	If ($iFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iFormat, $LOW_FIELD_FILENAME_FULL_PATH, $LOW_FIELD_FILENAME_NAME_AND_EXT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oFileNameField.FileFormat = $iFormat
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oFileNameField, $bOverwrite)

	If ($oFileNameField.IsFixed() = False) Then $oFileNameField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oFileNameField)
EndFunc   ;==>_LOWriter_FieldFileNameInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFileNameModify
; Description ...: Set or Retrieve a File Name Field's settings.
; Syntax ........: _LOWriter_FieldFileNameModify(Byref $oFileNameField[, $bIsFixed = Null[, $iFormat = Null]])
; Parameters ....: $oFileNameField      - [in/out] an object. A File Name field Object from a previous Insert or retrieval
;				   +									function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;				   $iFormat             - [optional] an integer value. Default is Null. The Data Format to  display. See
;				   +									Constants.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oFileNameField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $iFormat not an Integer, Less than 0, or greater than 3. See Constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iFormat
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Until at least L.O. Version 7.3.4.2, there is a bug where the wrong Path Format type is displayed when the
;						content is set to Fixed = True. For example, $LOW_FIELD_FILENAME_NAME_AND_EXT, displays in the format
;							of $LOW_FIELD_FILENAME_NAME.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;File Name Constants: $LOW_FIELD_FILENAME_FULL_PATH(0), The content of the URL is completely displayed.
;						$LOW_FIELD_FILENAME_PATH(1), Only the path of the file is displayed.
;						$LOW_FIELD_FILENAME_NAME(2), Only the name of the file without the file extension is displayed.
;						$LOW_FIELD_FILENAME_NAME_AND_EXT(3), The file name including the file extension is displayed.
; Related .......: _LOWriter_FieldFileNameInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFileNameModify(ByRef $oFileNameField, $bIsFixed = Null, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFileName[2]

	If Not IsObj($oFileNameField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iFormat, $bIsFixed) Then
		__LOWriter_ArrayFill($avFileName, $oFileNameField.IsFixed(), $oFileNameField.FileFormat())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avFileName)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oFileNameField.IsFixed = $bIsFixed
		$iError = ($oFileNameField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iFormat, $LOW_FIELD_FILENAME_FULL_PATH, $LOW_FIELD_FILENAME_NAME_AND_EXT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oFileNameField.FileFormat = $iFormat
		$iError = ($oFileNameField.FileFormat() = $iFormat) ? $iError : BitOR($iError, 1)
	EndIf

	If ($oFileNameField.IsFixed() = False) Then $oFileNameField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldFileNameModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncHiddenParInsert
; Description ...: Insert a Hidden Paragraph Field.
; Syntax ........: _LOWriter_FieldFuncHiddenParInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $sCondition = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $sCondition          - [optional] a string value. Default is Null. The condition to evaluate.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sCondition not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.HiddenParagraph" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Hidden Paragraph Field. Returning
;				   +											the Hidden Paragraph Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldFuncHiddenParModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncHiddenParInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sCondition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oHidParField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oHidParField = $oDoc.createInstance("com.sun.star.text.TextField.HiddenParagraph")
	If Not IsObj($oHidParField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oHidParField.Condition = $sCondition
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oHidParField, $bOverwrite)

	$oHidParField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oHidParField)
EndFunc   ;==>_LOWriter_FieldFuncHiddenParInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncHiddenParModify
; Description ...: Set or Retrieve a Hidden Paragraph Field's settings.
; Syntax ........: _LOWriter_FieldFuncHiddenParModify(Byref $oHidParField[, $sCondition = Null])
; Parameters ....: $oHidParField        - [in/out] an object. A Hidden Paragraph field Object from a previous Insert or retrieval
;				   +									function.
;                  $sCondition          - [optional] a string value. Default is Null. The condition to evaluate.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oHidParField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCondition not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1
;				   |								1 = Error setting $sCondition
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
;				   +								The second Element is a boolean whether the Paragraph is Hidden(True) or
;				   +								Visible(False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldFuncHiddenParInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncHiddenParModify(ByRef $oHidParField, $sCondition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHidPar[2]

	If Not IsObj($oHidParField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sCondition) Then
		__LOWriter_ArrayFill($avHidPar, $oHidParField.Condition(), ($oHidParField.IsHidden()) ? False : True) ;"IsHidden" Is Backwards
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avHidPar)
	EndIf

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oHidParField.Condition = $sCondition
		$iError = ($oHidParField.Condition() = $sCondition) ? $iError : BitOR($iError, 1)
	EndIf

	$oHidParField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldFuncHiddenParModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncHiddenTextInsert
; Description ...: Insert a Hidden Text Field.
; Syntax ........: _LOWriter_FieldFuncHiddenTextInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $sCondition = Null[, $sText = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $sCondition          - [optional] a string value. Default is Null. The Condition to evaluate.
;                  $sText               - [optional] a string value. Default is Null. The Text to show or hide.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sCondition not a String.
;				   @Error 1 @Extended 6 Return 0 = $sText not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.HiddenText" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Hidden Text Field. Returning
;				   +											the Hidden Text Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldFuncHiddenTextModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncHiddenTextInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sCondition = Null, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oHidTxtField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oHidTxtField = $oDoc.createInstance("com.sun.star.text.TextField.HiddenText")
	If Not IsObj($oHidTxtField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oHidTxtField.Condition = $sCondition
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oHidTxtField.Content = $sText
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oHidTxtField, $bOverwrite)

	$oHidTxtField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oHidTxtField)
EndFunc   ;==>_LOWriter_FieldFuncHiddenTextInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncHiddenTextModify
; Description ...: Set or Retrieve a Hidden Text Field's settings.
; Syntax ........: _LOWriter_FieldFuncHiddenTextModify(Byref $oHidTxtField[, $sCondition = Null[, $sText = Null]])
; Parameters ....: $oHidTxtField        - [in/out] an object. A Hidden Text field Object from a previous Insert or retrieval
;				   +									function.
;                  $sCondition          - [optional] a string value. Default is Null. The Condition to evaluate.
;                  $sText               - [optional] a string value. Default is Null. The Text to show or hide.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oHidTxtField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCondition not a String.
;				   @Error 1 @Extended 3 Return 0 = $sText not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $sCondition
;				   |								2 = Error setting $sText
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 3 Element Array with values in order of function parameters.
;				   +								The Third Element is a boolean whether the Text is Hidden(True) Or
;				   +								Visible(False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldFuncHiddenTextInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncHiddenTextModify(ByRef $oHidTxtField, $sCondition = Null, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHidPar[3]

	If Not IsObj($oHidTxtField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sCondition, $sText) Then
		__LOWriter_ArrayFill($avHidPar, $oHidTxtField.Condition(), $oHidTxtField.Content(), ($oHidTxtField.IsHidden()) ? False : True) ;"IsHidden" Is Backwards
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avHidPar)
	EndIf

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oHidTxtField.Condition = $sCondition
		$iError = ($oHidTxtField.Condition() = $sCondition) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oHidTxtField.Content = $sText
		$iError = ($oHidTxtField.Content() = $sText) ? $iError : BitOR($iError, 2)
	EndIf

	$oHidTxtField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldFuncHiddenTextModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncInputInsert
; Description ...: Insert a Input Field.
; Syntax ........: _LOWriter_FieldFuncInputInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $sReference = Null[, $sText = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $sReference          - [optional] a string value. Default is Null. The Reference to display for the input
;				   +								field.
;                  $sText               - [optional] a string value. Default is Null. The Text to insert in the Input Field.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sReference not a String.
;				   @Error 1 @Extended 6 Return 0 = $sText not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.Input" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Input Field. Returning
;				   +											the Input Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldFuncInputModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncInputInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sReference = Null, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oInputField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oInputField = $oDoc.createInstance("com.sun.star.text.TextField.Input")
	If Not IsObj($oInputField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($sReference <> Null) Then
		If Not IsString($sReference) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oInputField.Hint = $sReference
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oInputField.Content = $sText
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oInputField, $bOverwrite)

	$oInputField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oInputField)
EndFunc   ;==>_LOWriter_FieldFuncInputInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncInputModify
; Description ...: Set or Retrieve a Input Field's settings.
; Syntax ........: _LOWriter_FieldFuncInputModify(Byref $oInputField[, $sReference = Null[, $sText = Null]])
; Parameters ....: $oInputField         - [in/out] an object. A Input field Object from a previous Insert or retrieval
;				   +									function.
;                  $sReference          - [optional] a string value. Default is Null. The Reference to display for the input
;				   +								field.
;                  $sText               - [optional] a string value. Default is Null. The Text to insert in the Input Field.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oHidTxtField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sReference not a String.
;				   @Error 1 @Extended 3 Return 0 = $sText not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $sReference
;				   |								2 = Error setting $sText
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldFuncInputInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncInputModify(ByRef $oInputField, $sReference = Null, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asInput[2]

	If Not IsObj($oInputField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sReference, $sText) Then
		__LOWriter_ArrayFill($asInput, $oInputField.Hint(), $oInputField.Content())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $asInput)
	EndIf

	If ($sReference <> Null) Then
		If Not IsString($sReference) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oInputField.Hint = $sReference
		$iError = ($oInputField.Hint() = $sReference) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oInputField.Content = $sText
		$iError = ($oInputField.Content() = $sText) ? $iError : BitOR($iError, 2)
	EndIf

	$oInputField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldFuncInputModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncPlaceholderInsert
; Description ...: Insert a Placeholder Field.
; Syntax ........: _LOWriter_FieldFuncPlaceholderInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $iPHolderType = Null[, $sPHolderName = Null[, $sReference = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iPHolderType        - [optional] an integer value. Default is Null. The type of Placeholder to insert. See
;				   +									Constants.
;                  $sPHolderName        - [optional] a string value. Default is Null. The Placeholder's name.
;                  $sReference          - [optional] a string value. Default is Null. A Reference to display when the mouse
;				   +									hovers the Placeholder.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iPHolderType not an Integer, less than 0 or greater than 4. See Constants.
;				   @Error 1 @Extended 6 Return 0 = $sPHolderName not a String.
;				   @Error 1 @Extended 7 Return 0 = $sReference not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.JumpEdit" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Placeholder Field. Returning
;				   +											the Placeholder Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Placehold Type Constants: $LOW_FIELD_PLACEHOLD_TYPE_TEXT(0), The field represents a piece of text.
;							$LOW_FIELD_PLACEHOLD_TYPE_TABLE(1), The field initiates the insertion of a text table.
;							$LOW_FIELD_PLACEHOLD_TYPE_FRAME(2), The field initiates the insertion of a text frame.
;							$LOW_FIELD_PLACEHOLD_TYPE_GRAPHIC(3), The field initiates the insertion of a graphic object.
;							$LOW_FIELD_PLACEHOLD_TYPE_OBJECT(4), The field initiates the insertion of an embedded object.
; Related .......: _LOWriter_FieldFuncPlaceholderModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncPlaceholderInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $iPHolderType = Null, $sPHolderName = Null, $sReference = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPHolderField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oPHolderField = $oDoc.createInstance("com.sun.star.text.TextField.JumpEdit")
	If Not IsObj($oPHolderField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($iPHolderType <> Null) Then
		If Not __LOWriter_IntIsBetween($iPHolderType, $LOW_FIELD_PLACEHOLD_TYPE_TEXT, $LOW_FIELD_PLACEHOLD_TYPE_OBJECT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oPHolderField.PlaceHolderType = $iPHolderType
	EndIf

	If ($sPHolderName <> Null) Then
		If Not IsString($sPHolderName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oPHolderField.PlaceHolder = $sPHolderName
	EndIf

	If ($sReference <> Null) Then
		If Not IsString($sReference) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oPHolderField.Hint = $sReference
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oPHolderField, $bOverwrite)

	$oPHolderField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oPHolderField)
EndFunc   ;==>_LOWriter_FieldFuncPlaceholderInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncPlaceholderModify
; Description ...: Set or Retrieve a Placeholder Field's settings.
; Syntax ........: _LOWriter_FieldFuncPlaceholderModify(Byref $oPHolderField[, $iPHolderType = Null[, $sPHolderName = Null[, $sReference = Null]]])
; Parameters ....: $oPHolderField       - [in/out] an object. A Placeholder field Object from a previous Insert or retrieval
;				   +									function.
;                  $iPHolderType        - [optional] an integer value. Default is Null. The type of Placeholder to insert. See
;				   +									Constants.
;                  $sPHolderName        - [optional] a string value. Default is Null. The Placeholder's name.
;                  $sReference          - [optional] a string value. Default is Null. A Reference to display when the mouse
;				   +									hovers the Placeholder.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oPHolderField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iPHolderType not an Integer, less than 0 or greater than 4. See Constants.
;				   @Error 1 @Extended 3 Return 0 = $sPHolderName not a String.
;				   @Error 1 @Extended 4 Return 0 = $sReference not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $iPHolderType
;				   |								2 = Error setting $sPHolderName
;				   |								4 = Error setting $sReference
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;Placehold Type Constants: $LOW_FIELD_PLACEHOLD_TYPE_TEXT(0), The field represents a piece of text.
;							$LOW_FIELD_PLACEHOLD_TYPE_TABLE(1), The field initiates the insertion of a text table.
;							$LOW_FIELD_PLACEHOLD_TYPE_FRAME(2), The field initiates the insertion of a text frame.
;							$LOW_FIELD_PLACEHOLD_TYPE_GRAPHIC(3), The field initiates the insertion of a graphic object.
;							$LOW_FIELD_PLACEHOLD_TYPE_OBJECT(4), The field initiates the insertion of an embedded object.
; Related .......: _LOWriter_FieldFuncPlaceholderInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncPlaceholderModify(ByRef $oPHolderField, $iPHolderType = Null, $sPHolderName = Null, $sReference = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asPHolder[3]

	If Not IsObj($oPHolderField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iPHolderType, $sPHolderName, $sReference) Then
		__LOWriter_ArrayFill($asPHolder, $oPHolderField.PlaceHolderType(), $oPHolderField.PlaceHolder(), $oPHolderField.Hint())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $asPHolder)
	EndIf

	If ($iPHolderType <> Null) Then
		If Not __LOWriter_IntIsBetween($iPHolderType, $LOW_FIELD_PLACEHOLD_TYPE_TEXT, $LOW_FIELD_PLACEHOLD_TYPE_OBJECT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oPHolderField.PlaceHolderType = $iPHolderType
		$iError = ($oPHolderField.PlaceHolderType() = $iPHolderType) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sPHolderName <> Null) Then
		If Not IsString($sPHolderName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oPHolderField.PlaceHolder = $sPHolderName
		$iError = ($oPHolderField.PlaceHolder() = $sPHolderName) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sReference <> Null) Then
		If Not IsString($sReference) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oPHolderField.Hint = $sReference
		$iError = ($oPHolderField.Hint() = $sReference) ? $iError : BitOR($iError, 4)
	EndIf

	$oPHolderField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldFuncPlaceholderModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldGetAnchor
; Description ...: Retrieve the Anchor Cursor Object for a Field inserted in a document.
; Syntax ........: _LOWriter_FieldGetAnchor(Byref $oField)
; Parameters ....: $oField              - [in/out] an object. A Field Object returned from a previous Insert or Retrieve
;				   +						function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oField not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Field anchor Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning requested Field Anchor Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldsGetList, _LOWriter_FieldsAdvGetList, _LOWriter_FieldsDocInfoGetList, _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldGetAnchor(ByRef $oField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFieldAnchor

	If Not IsObj($oField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oFieldAnchor = $oField.Anchor.Text.createTextCursorByRange($oField.Anchor())
	If Not IsObj($oFieldAnchor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oFieldAnchor)
EndFunc   ;==>_LOWriter_FieldGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldInputListInsert
; Description ...: Insert a Input List Field.
; Syntax ........: _LOWriter_FieldInputListInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $asItems = Null[, $sName = Null[, $sSelectedItem = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $asItems             - [optional] an array of strings. Default is Null. A single column Array of Strings to
;				   +									colonize the List with.
;                  $sName               - [optional] a string value. Default is Null. The name of the Input List Field.
;                  $sSelectedItem       - [optional] a string value. Default is Null. The Item in the list to be currently
;				   +									selected. Defaults to "" if Item is not found.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $asItems not an Array.
;				   @Error 1 @Extended 6 Return 0 = $sName not a String.
;				   @Error 1 @Extended 7 Return 0 = $sSelectedItem not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.DropDown" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Input List field, returning
;				   +										Input List Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldInputListModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldInputListInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $asItems = Null, $sName = Null, $sSelectedItem = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oInputField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oInputField = $oDoc.createInstance("com.sun.star.text.TextField.DropDown")
	If Not IsObj($oInputField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($asItems <> Null) Then
		If Not IsArray($asItems) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oInputField.Items = $asItems
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oInputField.Name = $sName
	EndIf

	If ($sSelectedItem <> Null) Then
		If Not IsString($sSelectedItem) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oInputField.SelectedItem = $sSelectedItem
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oInputField, $bOverwrite)

	$oInputField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oInputField)
EndFunc   ;==>_LOWriter_FieldInputListInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldInputListModify
; Description ...: Set or Retrieve a Input List Field's settings.
; Syntax ........: _LOWriter_FieldInputListModify(Byref $oInputField[, $asItems = Null[, $sName = Null[, $sSelectedItem = Null]]])
; Parameters ....: $oInputField         - [in/out] an object. A Input List field Object from a previous Insert or retrieval
;				   +									function.
;                  $asItems             - [optional] an array of strings. Default is Null. A single column Array of Strings to
;				   +									colonize the List with.
;                  $sName               - [optional] a string value. Default is Null. The name of the Input List Field.
;                  $sSelectedItem       - [optional] a string value. Default is Null. The Item in the list to be currently
;				   +									selected. Defaults to "" if Item is not found.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oInputField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $asItems not an Array.
;				   @Error 1 @Extended 3 Return 0 = $sName not a String.
;				   @Error 1 @Extended 4 Return 0 = $sSelectedItem not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $asItems
;				   |								2 = Error setting $sName
;				   |								4 = Error setting $sSelectedItem
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldInputListInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldInputListModify(ByRef $oInputField, $asItems = Null, $sName = Null, $sSelectedItem = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDropDwn[3]

	If Not IsObj($oInputField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($asItems, $sName, $sSelectedItem) Then
		__LOWriter_ArrayFill($avDropDwn, $oInputField.Items(), $oInputField.Name(), $oInputField.SelectedItem())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDropDwn)
	EndIf

	If ($asItems <> Null) Then
		If Not IsArray($asItems) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oInputField.Items = $asItems
		$iError = (UBound($oInputField.Items()) = UBound($asItems)) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oInputField.Name = $sName
		$iError = ($oInputField.Name() = $sName) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sSelectedItem <> Null) Then
		If Not IsString($sSelectedItem) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oInputField.SelectedItem = $sSelectedItem
		$iError = ($oInputField.SelectedItem() = $sSelectedItem) ? $iError : BitOR($iError, 4)
	EndIf

	$oInputField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldInputListModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldPageNumberInsert
; Description ...: Insert a Page number field.
; Syntax ........: _LOWriter_FieldPageNumberInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $iNumFormat = Null[, $iOffset = Null[, $iPageNumType = Null[, $sUserText = Null]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iNumFormat          - [optional] an integer value. Default is Null. The numbering format to use for Page
;				   +						numbering. See Constants.
;                  $iOffset             - [optional] an integer value. Default is Null. The number of pages to minus or add to
;				   +									the page Number.
;                  $iPageNumType        - [optional] an integer value. Default is Null. The Page Number type, either previous,
;				   +									current or next page. See Constants.
;                  $sUserText           - [optional] a string value. Default is Null. The custom User text to display. Only valid
;				   +									if $iNumFormat is set to $LOW_NUM_STYLE_CHAR_SPECIAL(6).
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iNumFormat not an Integer, less than 0 or greater than 71. See Constants.
;				   @Error 1 @Extended 6 Return 0 = $iOffset not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iPageNumType not an Integer, less than 0 or greater than 2. See Constants.
;				   @Error 1 @Extended 8 Return 0 = $sUserText not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.PageNumber" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Page Number field, returning Page Num.
;				   +										Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Numbering Format Constants: $LOW_NUM_STYLE_CHARS_UPPER_LETTER(0), Numbering is put in upper case letters. ("A, B, C, D)
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
;Page Number Type Constants: $LOW_PAGE_NUM_TYPE_PREV(0), The Previous Page's page number.
;								$LOW_PAGE_NUM_TYPE_CURRENT(1), The current page number.
;								$LOW_PAGE_NUM_TYPE_NEXT(2), The Next Page's page number.
; Related .......: _LOWriter_FieldPageNumberModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldPageNumberInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $iNumFormat = Null, $iOffset = Null, $iPageNumType = Null, $sUserText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oPageField = $oDoc.createInstance("com.sun.star.text.TextField.PageNumber")
	If Not IsObj($oPageField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oPageField.NumberingType = $iNumFormat
	Else
		$oPageField.NumberingType = $LOW_NUM_STYLE_PAGE_DESCRIPTOR
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oPageField.Offset = $iOffset
	EndIf

	If ($iPageNumType <> Null) Then
		If Not __LOWriter_IntIsBetween($iPageNumType, $LOW_PAGE_NUM_TYPE_PREV, $LOW_PAGE_NUM_TYPE_NEXT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oPageField.SubType = $iPageNumType

		If ($iPageNumType = $LOW_PAGE_NUM_TYPE_PREV) Then
			$oPageField.Offset = ($oPageField.Offset() - 1) ;If SubType is Set to Prev. Set offset to minus 1 of current value
		ElseIf ($iPageNumType = $LOW_PAGE_NUM_TYPE_NEXT) Then
			$oPageField.Offset = ($oPageField.Offset() + 1) ;If SubType is Set to Next. Set offset to plus 1 of current value
		EndIf
	Else
		$oPageField.SubType = $LOW_PAGE_NUM_TYPE_CURRENT ;If not set, page number Sub Type is auto set to Prev. Instead of current.
	EndIf

	If ($sUserText <> Null) Then
		If Not IsString($sUserText) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oPageField.UserText = $sUserText
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oPageField, $bOverwrite)

	$oPageField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oPageField)
EndFunc   ;==>_LOWriter_FieldPageNumberInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldPageNumberModify
; Description ...: Set or Retrieve Page NUmber Field settings.
; Syntax ........: _LOWriter_FieldPageNumberModify(Byref $oDoc, Byref $oPageNumField[, $iNumFormat = Null[, $iOffset = Null[, $iPageNumType = Null[, $sUserText = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oPageNumField       - [in/out] an object. A Page Number field Object from a previous Insert or retrieval
;				   +									function.
;                  $iNumFormat          - [optional] an integer value. Default is Null. The numbering format to use for Page
;				   +						numbering. See Constants.
;                  $iOffset             - [optional] an integer value. Default is Null. The number of pages to minus or add to
;				   +									the page Number.
;                  $iPageNumType        - [optional] an integer value. Default is Null. The Page Number type, either previous,
;				   +									current or next page. See Constants.
;                  $sUserText           - [optional] a string value. Default is Null. The custom User text to display. Only valid
;				   +									if $iNumFormat is set to $LOW_NUM_STYLE_CHAR_SPECIAL(6).
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oPageNumField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iNumFormat not an Integer, less than 0 or greater than 71. See Constants.
;				   @Error 1 @Extended 4 Return 0 = $iOffset not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iPageNumType not an Integer, less than 0 or greater than 2. See Constants.
;				   @Error 1 @Extended 6 Return 0 = $sUserText not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.PageNumber" Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $iNumFormat
;				   |								2 = Error setting $iOffset
;				   |								4 = Error setting $iPageNumType
;				   |								8 = Error setting $sUserText
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;Numbering Format Constants: $LOW_NUM_STYLE_CHARS_UPPER_LETTER(0), Numbering is put in upper case letters. ("A, B, C, D)
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
;Page Number Type Constants: $LOW_PAGE_NUM_TYPE_PREV(0), The Previous Page's page number.
;								$LOW_PAGE_NUM_TYPE_CURRENT(1), The current page number.
;								$LOW_PAGE_NUM_TYPE_NEXT(2), The Next Page's page number.
; Related .......: _LOWriter_FieldPageNumberInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldPageNumberModify(ByRef $oDoc, ByRef $oPageNumField, $iNumFormat = Null, $iOffset = Null, $iPageNumType = Null, $sUserText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avField[4]
	Local $iError = 0
	Local $oNewPageNumField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageNumField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($iNumFormat, $iOffset, $iPageNumType, $sUserText) Then
		__LOWriter_ArrayFill($avField, $oPageNumField.NumberingType(), $oPageNumField.Offset(), $oPageNumField.SubType(), $oPageNumField.UserText())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avField)
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

		$oNewPageNumField = $oDoc.createInstance("com.sun.star.text.TextField.PageNumber")
		If Not IsObj($oNewPageNumField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

		;It doesn't work to just set a new Numbering type for an already inserted Page Number, so I have to create a new one and
		;then insert it.
		With $oNewPageNumField
			.NumberingType = $iNumFormat
			.Offset = $oPageNumField.Offset()
			.SubType = $oPageNumField.SubType()
			.UserText = $oPageNumField.UserText()
		EndWith

		$oDoc.Text.createTextCursorByRange($oPageNumField.Anchor()).Text.insertTextContent($oPageNumField.Anchor(), $oNewPageNumField, True)

		;Update the Old Page nUmber Field Object to the new one.
		$oPageNumField = $oNewPageNumField

		$iError = ($oPageNumField.NumberingType() = $iNumFormat) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oPageNumField.Offset = $iOffset
		$iError = ($oPageNumField.Offset() = $iOffset) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iPageNumType <> Null) Then
		If Not __LOWriter_IntIsBetween($iPageNumType, $LOW_PAGE_NUM_TYPE_PREV, $LOW_PAGE_NUM_TYPE_NEXT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oPageNumField.SubType = $iPageNumType
		$iError = ($oPageNumField.SubType() = $iPageNumType) ? $iError : BitOR($iError, 4)
	EndIf

	If ($sUserText <> Null) Then
		If Not IsString($sUserText) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oPageNumField.UserText = $sUserText
		$iError = ($oPageNumField.UserText() = $sUserText) ? $iError : BitOR($iError, 8)
	EndIf

	$oPageNumField.Update()



	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldPageNumberModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefBookMarkInsert
; Description ...: Insert a Bookmark Reference Field.
; Syntax ........: _LOWriter_FieldRefBookMarkInsert(Byref $oDoc, Byref $oCursor, $sBookmarkName[, $bOverwrite = False[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $sBookmarkName       - a string value. The Bookmark name to Reference.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iRefUsing            - [optional] an integer value. Default is Null. The Type of reference to use to
;				   +									reference the bookmark, see Constants.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $sBookmarkName not a String.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = Document does not contain a Bookmark by the same name as called in
;				   +									$sBookmarkName.
;				   @Error 1 @Extended 7 Return 0 = $iRefUsing not an Integer, Less than 0 or greater than 4. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.GetReference" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Bookmark Reference Field. Returning
;				   +											the Bookmark Reference Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Refer Using Constants: $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED(0), The page number is displayed using Arabic numbers.
;							$LOW_FIELD_REF_USING_CHAPTER(1), The number of the chapter is displayed.
;							$LOW_FIELD_REF_USING_REF_TEXT(2), The reference text is displayed.
;							$LOW_FIELD_REF_USING_ABOVE_BELOW(3), The reference is displayed as one of the words, "above" or
;								"below".
;							$LOW_FIELD_REF_USING_PAGE_NUM_STYLED(4), The page number is displayed using the numbering type
;								defined in the page style of the reference position.
; Related .......: _LOWriter_FieldRefBookMarkModify, _LOWriter_DocBookmarkInsert, _LOWriter_DocBookmarksList,
;					 _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefBookMarkInsert(ByRef $oDoc, ByRef $oCursor, $sBookmarkName, $bOverwrite = False, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBookmarkRefField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sBookmarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	If Not _LOWriter_DocBookmarksHasName($oDoc, $sBookmarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$oBookmarkRefField = $oDoc.createInstance("com.sun.star.text.TextField.GetReference")
	If Not IsObj($oBookmarkRefField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$oBookmarkRefField.SourceName = $sBookmarkName
	$oBookmarkRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_BOOKMARK

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oBookmarkRefField.ReferenceFieldPart = $iRefUsing
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oBookmarkRefField, $bOverwrite)

	$oBookmarkRefField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oBookmarkRefField)
EndFunc   ;==>_LOWriter_FieldRefBookMarkInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefBookMarkModify
; Description ...: Set or Retrieve a Bookmark Reference Field's settings.
; Syntax ........: _LOWriter_FieldRefBookMarkModify(Byref $oDoc, Byref $oBookmarkRefField[, $sBookmarkName = Null[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oBookmarkRefField   - [in/out] an object. A Bookmark Reference field Object from a previous Insert or
;				   +									retrieval function.
;                  $sBookmarkName       - [optional] a string value. Default is Null. The Bookmark name to Reference.
;                  $iRefUsing            - [optional] an integer value. Default is Null. The Type of reference to use to
;				   +									reference the bookmark, see Constants.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oBookmarkRefField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sBookmarkName not a String.
;				   @Error 1 @Extended 4 Return 0 = Document does not contain a Bookmark by the same name as called in
;				   +									$sBookmarkName.
;				   @Error 1 @Extended 5 Return 0 = $iRefUsing not an Integer, Less than 0 or greater than 4. See Constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $sBookmarkName
;				   |								2 = Error setting $iRefUsing
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;Refer Using Constants: $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED(0), The page number is displayed using Arabic numbers.
;							$LOW_FIELD_REF_USING_CHAPTER(1), The number of the chapter is displayed.
;							$LOW_FIELD_REF_USING_REF_TEXT(2), The reference text is displayed.
;							$LOW_FIELD_REF_USING_ABOVE_BELOW(3), The reference is displayed as one of the words, "above" or
;								"below".
;							$LOW_FIELD_REF_USING_PAGE_NUM_STYLED(4), The page number is displayed using the numbering type
;								defined in the page style of the reference position.
; Related .......: _LOWriter_FieldRefBookMarkInsert, _LOWriter_DocBookmarkInsert, _LOWriter_DocBookmarksList,
;					_LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefBookMarkModify(ByRef $oDoc, ByRef $oBookmarkRefField, $sBookmarkName = Null, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avBook[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oBookmarkRefField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sBookmarkName, $iRefUsing) Then
		__LOWriter_ArrayFill($avBook, $oBookmarkRefField.SourceName(), $oBookmarkRefField.ReferenceFieldPart())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avBook)
	EndIf

	If ($sBookmarkName <> Null) Then
		If Not IsString($sBookmarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		If Not _LOWriter_DocBookmarksHasName($oDoc, $sBookmarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oBookmarkRefField.SourceName = $sBookmarkName
		$oBookmarkRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_BOOKMARK ;Set Type to Bookmark in case input field Obj is a diff type.
		$iError = ($oBookmarkRefField.SourceName = $sBookmarkName) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oBookmarkRefField.ReferenceFieldPart = $iRefUsing
		$iError = ($oBookmarkRefField.ReferenceFieldPart = $iRefUsing) ? $iError : BitOR($iError, 2)
	EndIf

	$oBookmarkRefField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldRefBookMarkModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefEndnoteInsert
; Description ...: Insert a Endnote Reference Field.
; Syntax ........: _LOWriter_FieldRefEndnoteInsert(Byref $oDoc, Byref $oCursor, Byref $oEndNote[, $bOverwrite = False[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $oEndNote            - [in/out] an object. The Endnote Object returned from a previous Insert or Retrieve
;				   +						function, to reference.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iRefUsing           - [optional] an integer value. Default is Null. The Type of reference to use to
;				   +									reference the Endnote, see Constants.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $oEndNote not an Object.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iRefUsing not an Integer, Less than 0 or greater than 4. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.GetReference" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Endnote Reference Field. Returning
;				   +											the Endnote Reference Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Refer Using Constants: $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED(0), The page number is displayed using Arabic numbers.
;							$LOW_FIELD_REF_USING_CHAPTER(1), The number of the chapter is displayed.
;							$LOW_FIELD_REF_USING_REF_TEXT(2), The reference text is displayed.
;							$LOW_FIELD_REF_USING_ABOVE_BELOW(3), The reference is displayed as one of the words, "above" or
;								"below".
;							$LOW_FIELD_REF_USING_PAGE_NUM_STYLED(4), The page number is displayed using the numbering type
;								defined in the page style of the reference position.
; Related .......: _LOWriter_FieldRefEndnoteModify, _LOWriter_EndnoteInsert, _LOWriter_EndnotesGetList,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefEndnoteInsert(ByRef $oDoc, ByRef $oCursor, ByRef $oEndNote, $bOverwrite = False, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oENoteRefField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$oENoteRefField = $oDoc.createInstance("com.sun.star.text.TextField.GetReference")
	If Not IsObj($oENoteRefField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$oENoteRefField.SourceName = ""
	$oENoteRefField.SequenceNumber = $oEndNote.ReferenceId()
	$oENoteRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_ENDNOTE

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oENoteRefField.ReferenceFieldPart = $iRefUsing
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oENoteRefField, $bOverwrite)

	$oENoteRefField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oENoteRefField)
EndFunc   ;==>_LOWriter_FieldRefEndnoteInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefEndnoteModify
; Description ...: Set or Retrieve a Endnote Reference Field's settings.
; Syntax ........: _LOWriter_FieldRefEndnoteModify(Byref $oDoc, Byref $oEndNoteRefField[, $oEndNote = Null[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oEndNoteRefField    - [in/out] an object. A Endnote Reference field Object from a previous Insert or
;				   +									retrieval function.
;                  $oEndNote            - [optional] an object. Default is Null. The Endnote Object returned from a previous
;				   +						Insert or Retrieve function, to reference.
;                  $iRefUsing           - [optional] an integer value. Default is Null. The Type of reference to use to
;				   +									reference the Endnote, see Constants.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oEndNoteRefField not an Object.
;				   @Error 1 @Extended 3 Return 0 = Optional Parameters set to null, but $oEndNoteRefField object is not a
;				   +									listed as a Endnote Reference type field.
;				   @Error 1 @Extended 4 Return 0 = $oEndNote not an Object.
;				   @Error 1 @Extended 5 Return 0 = $iRefUsing not an Integer, Less than 0 or greater than 4. See Constants.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Endnote Object for setting return.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $oEndNote
;				   |								2 = Error setting $iRefUsing
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;Refer Using Constants: $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED(0), The page number is displayed using Arabic numbers.
;							$LOW_FIELD_REF_USING_CHAPTER(1), The number of the chapter is displayed.
;							$LOW_FIELD_REF_USING_REF_TEXT(2), The reference text is displayed.
;							$LOW_FIELD_REF_USING_ABOVE_BELOW(3), The reference is displayed as one of the words, "above" or
;								"below".
;							$LOW_FIELD_REF_USING_PAGE_NUM_STYLED(4), The page number is displayed using the numbering type
;								defined in the page style of the reference position.
; Related .......: _LOWriter_FieldRefEndnoteInsert, _LOWriter_EndnoteInsert, _LOWriter_EndnotesGetList, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefEndnoteModify(ByRef $oDoc, ByRef $oEndNoteRefField, $oEndNote = Null, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iSourceSeq
	Local $avFoot[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oEndNoteRefField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($oEndNote, $iRefUsing) Then
		If Not ($oEndNoteRefField.ReferenceFieldSource() = $LOW_FIELD_REF_TYPE_ENDNOTE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		If $oDoc.Endnotes.hasElements() Then
			$iSourceSeq = $oEndNoteRefField.SequenceNumber()
			For $i = 0 To $oDoc.Endnotes.Count() - 1 ;Locate referenced Endnote.
				If ($oDoc.Endnotes.getByIndex($i).ReferenceId() = $iSourceSeq) Then
					__LOWriter_ArrayFill($avFoot, $oDoc.Endnotes.getByIndex($i), $oEndNoteRefField.ReferenceFieldPart())
					Return SetError($__LOW_STATUS_SUCCESS, 1, $avFoot)
				EndIf
				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
			Next

		EndIf
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ;Error retrieving ENote Obj
	EndIf

	If ($oEndNote <> Null) Then
		If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oEndNoteRefField.SourceName = ""
		$oEndNoteRefField.SequenceNumber = $oEndNote.ReferenceId()
		$oEndNoteRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_ENDNOTE ;Set Type to Endnote in case input field Obj is a diff type.
		$iError = ($oEndNoteRefField.SequenceNumber = $oEndNote.ReferenceId()) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oEndNoteRefField.ReferenceFieldPart = $iRefUsing
		$iError = ($oEndNoteRefField.ReferenceFieldPart = $iRefUsing) ? $iError : BitOR($iError, 2)
	EndIf

	$oEndNoteRefField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldRefEndnoteModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefFootnoteInsert
; Description ...: Insert a Footnote Reference Field.
; Syntax ........: _LOWriter_FieldRefFootnoteInsert(Byref $oDoc, Byref $oCursor, Byref $oFootNote[, $bOverwrite = False[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $oFootNote           - [in/out] an object. The Footnote Object returned from a previous Insert or Retrieve
;				   +						function, to reference.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iRefUsing            - [optional] an integer value. Default is Null. The Type of reference to use to
;				   +									reference the Footnote, see Constants.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $oFootNote not an Object.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iRefUsing not an Integer, Less than 0 or greater than 4. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.GetReference" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Footnote Reference Field. Returning
;				   +											the Footnote Reference Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Refer Using Constants: $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED(0), The page number is displayed using Arabic numbers.
;							$LOW_FIELD_REF_USING_CHAPTER(1), The number of the chapter is displayed.
;							$LOW_FIELD_REF_USING_REF_TEXT(2), The reference text is displayed.
;							$LOW_FIELD_REF_USING_ABOVE_BELOW(3), The reference is displayed as one of the words, "above" or
;								"below".
;							$LOW_FIELD_REF_USING_PAGE_NUM_STYLED(4), The page number is displayed using the numbering type
;								defined in the page style of the reference position.
; Related .......: _LOWriter_FieldRefFootnoteModify, _LOWriter_FootnoteInsert, _LOWriter_FootnotesGetList,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefFootnoteInsert(ByRef $oDoc, ByRef $oCursor, ByRef $oFootNote, $bOverwrite = False, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFNoteRefField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$oFNoteRefField = $oDoc.createInstance("com.sun.star.text.TextField.GetReference")
	If Not IsObj($oFNoteRefField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$oFNoteRefField.SourceName = ""
	$oFNoteRefField.SequenceNumber = $oFootNote.ReferenceId()
	$oFNoteRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_FOOTNOTE

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oFNoteRefField.ReferenceFieldPart = $iRefUsing
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oFNoteRefField, $bOverwrite)

	$oFNoteRefField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oFNoteRefField)
EndFunc   ;==>_LOWriter_FieldRefFootnoteInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefFootnoteModify
; Description ...: Set or Retrieve a Footnote Reference Field's settings.
; Syntax ........: _LOWriter_FieldRefFootnoteModify(Byref $oDoc, Byref $oFootNoteRefField[, $oFootNote = Null[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oFootNoteRefField   - [in/out] an object. A Footnote Reference field Object from a previous Insert or
;				   +									retrieval function.
;                  $oFootNote           - [optional] an object. Default is Null. The Footnote Object returned from a previous
;				   +						Insert or Retrieve function, to reference.
;                  $iRefUsing           - [optional] an integer value. Default is Null. The Type of reference to use to
;				   +									reference the Footnote, see Constants.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oFootNoteRefField not an Object.
;				   @Error 1 @Extended 3 Return 0 = Optional Parameters set to null, but $oFootNoteRefField object is not a
;				   +									listed as a Footnote Reference type field.
;				   @Error 1 @Extended 4 Return 0 = $oFootNote not an Object.
;				   @Error 1 @Extended 5 Return 0 = $iRefUsing not an Integer, Less than 0 or greater than 4. See Constants.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Footnote Object for setting return.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $oFootNote
;				   |								2 = Error setting $iRefUsing
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;Refer Using Constants: $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED(0), The page number is displayed using Arabic numbers.
;							$LOW_FIELD_REF_USING_CHAPTER(1), The number of the chapter is displayed.
;							$LOW_FIELD_REF_USING_REF_TEXT(2), The reference text is displayed.
;							$LOW_FIELD_REF_USING_ABOVE_BELOW(3), The reference is displayed as one of the words, "above" or
;								"below".
;							$LOW_FIELD_REF_USING_PAGE_NUM_STYLED(4), The page number is displayed using the numbering type
;								defined in the page style of the reference position.
; Related .......: _LOWriter_FieldRefFootnoteInsert, _LOWriter_FootnoteInsert, _LOWriter_FootnotesGetList,
;					_LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefFootnoteModify(ByRef $oDoc, ByRef $oFootNoteRefField, $oFootNote = Null, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iSourceSeq
	Local $avFoot[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFootNoteRefField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($oFootNote, $iRefUsing) Then
		If Not ($oFootNoteRefField.ReferenceFieldSource() = $LOW_FIELD_REF_TYPE_FOOTNOTE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		If $oDoc.Footnotes.hasElements() Then
			$iSourceSeq = $oFootNoteRefField.SequenceNumber()
			For $i = 0 To $oDoc.Footnotes.Count() - 1 ;Locate referenced Footnote.
				If ($oDoc.Footnotes.getByIndex($i).ReferenceId() = $iSourceSeq) Then
					__LOWriter_ArrayFill($avFoot, $oDoc.Footnotes.getByIndex($i), $oFootNoteRefField.ReferenceFieldPart())
					Return SetError($__LOW_STATUS_SUCCESS, 1, $avFoot)
				EndIf
				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
			Next

		EndIf
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ;Error retrieving FNote Obj
	EndIf

	If ($oFootNote <> Null) Then
		If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oFootNoteRefField.SourceName = ""
		$oFootNoteRefField.SequenceNumber = $oFootNote.ReferenceId()
		$oFootNoteRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_FOOTNOTE ;Set Type to Footnote in case input field Obj is a diff type.
		$iError = ($oFootNoteRefField.SequenceNumber = $oFootNote.ReferenceId()) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oFootNoteRefField.ReferenceFieldPart = $iRefUsing
		$iError = ($oFootNoteRefField.ReferenceFieldPart = $iRefUsing) ? $iError : BitOR($iError, 2)
	EndIf

	$oFootNoteRefField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldRefFootnoteModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefGetType
; Description ...: Retrieve the type of Data a Reference Field is Referencing.
; Syntax ........: _LOWriter_FieldRefGetType(Byref $oRefField)
; Parameters ....: $oRefField           - [in/out] an object. a Reference Field Object from a previous Insert or Retrieve
;				   +						function.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRefField not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Returning the Data Type Source for the reference Field.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Reference Field can be referencing multiple different types of Data, such as a Reference Mark, or Bookmark,
;					etc.
;Reference Type Constants: $LOW_FIELD_REF_TYPE_REF_MARK(0), The source is referencing a reference mark.
;							$LOW_FIELD_REF_TYPE_SEQ_FIELD(1), The source is referencing a number sequence field. Such as a
;															Number range variable, numbered Paragraph, etc.
;							$LOW_FIELD_REF_TYPE_BOOKMARK(2), The source is referencing a bookmark.
;							$LOW_FIELD_REF_TYPE_FOOTNOTE(3), The source is referencing a footnote.
;							$LOW_FIELD_REF_TYPE_ENDNOTE(4), The source is referencing an endnote.
; Related .......: _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefGetType(ByRef $oRefField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oRefField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oRefField.ReferenceFieldSource())
EndFunc   ;==>_LOWriter_FieldRefGetType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefInsert
; Description ...: Insert a Reference Field.
; Syntax ........: _LOWriter_FieldRefInsert(Byref $oDoc, Byref $oCursor, $sRefMarkName[, $bOverwrite = False[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $sRefMarkName        - a string value. The Reference Mark Name to Reference.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iRefUsing           - [optional] an integer value. Default is Null. The Type of reference to insert, see
;				   +									Constants.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $sRefMarkName not a String.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = Document does not contain a Reference Mark by the same name as called in
;				   +									$sRefMarkName.
;				   @Error 1 @Extended 7 Return 0 = $iRefUsing not an Integer, Less than 0 or greater than 4. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Reference Marks Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.text.TextField.GetReference" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Reference Field. Returning the
;				   +											Reference Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Refer Using Constants: $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED(0), The page number is displayed using Arabic numbers.
;							$LOW_FIELD_REF_USING_CHAPTER(1), The number of the chapter is displayed.
;							$LOW_FIELD_REF_USING_REF_TEXT(2), The reference text is displayed.
;							$LOW_FIELD_REF_USING_ABOVE_BELOW(3), The reference is displayed as one of the words, "above" or
;								"below".
;							$LOW_FIELD_REF_USING_PAGE_NUM_STYLED(4), The page number is displayed using the numbering type
;								defined in the page style of the reference position.
; Related .......: _LOWriter_FieldRefModify, _LOWriter_FieldRefMarkSet, _LOWriter_FieldRefMarkList,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefInsert(ByRef $oDoc, ByRef $oCursor, $sRefMarkName, $bOverwrite = False, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMarks, $oMarkRefField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sRefMarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$oRefMarks = $oDoc.getReferenceMarks()
	If Not IsObj($oRefMarks) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If Not $oRefMarks.hasByName($sRefMarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$oMarkRefField = $oDoc.createInstance("com.sun.star.text.TextField.GetReference")
	If Not IsObj($oMarkRefField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$oMarkRefField.SourceName = $sRefMarkName
	$oMarkRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_REF_MARK

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oMarkRefField.ReferenceFieldPart = $iRefUsing
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oMarkRefField, $bOverwrite)

	$oMarkRefField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oMarkRefField)
EndFunc   ;==>_LOWriter_FieldRefInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefMarkDelete
; Description ...: Delete a Reference Mark by name.
; Syntax ........: _LOWriter_FieldRefMarkDelete(Byref $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sName               - a string value. The Reference Mark name to delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sName not a String.
;				   @Error 1 @Extended 3 Return 0 = Document does not contain a Reference Mark named the same as called in
;				   +								$sName
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Reference Marks Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Reference Mark object called in $sName.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to delete Reference Mark, but document still contains a
;				   +									Reference Mark by that name.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested Reference Mark.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefMarkSet, _LOWriter_FieldRefMarkList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefMarkDelete(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMark, $oRefMarks

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oRefMarks = $oDoc.getReferenceMarks()
	If Not IsObj($oRefMarks) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If Not $oRefMarks.hasByName($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$oRefMark = $oRefMarks.getByName($sName)
	If Not IsObj($oRefMark) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$oRefMark.dispose()

	Return ($oRefMarks.hasByName($sName)) ? SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldRefMarkDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefMarkGetAnchor
; Description ...: Retrieve the Anchor Cursor Object of a Reference Mark by Name.
; Syntax ........: _LOWriter_FieldRefMarkGetAnchor(Byref $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sName               - a string value. The Reference Mark name to retrieve the anchor for.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sName not a String.
;				   @Error 1 @Extended 3 Return 0 = Document does not contain a Reference Mark named the same as called in
;				   +								$sName
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Reference Marks Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Reference Mark object called in $sName.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning requested Reference Mark Anchor Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefMarkList, _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefMarkGetAnchor(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMark, $oRefMarks

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oRefMarks = $oDoc.getReferenceMarks()
	If Not IsObj($oRefMarks) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If Not $oRefMarks.hasByName($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$oRefMark = $oRefMarks.getByName($sName)
	If Not IsObj($oRefMark) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oRefMark.Anchor.Text.createTextCursorByRange($oRefMark.Anchor()))
EndFunc   ;==>_LOWriter_FieldRefMarkGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefMarkList
; Description ...: Retrieve an Array of Reference Mark names.
; Syntax ........: _LOWriter_FieldRefMarkList(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Reference Marks Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Array of Reference Mark Names.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully searched for Reference Marks, but document does
;				   +											not contain any.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for Reference Marks, returning Array of
;				   +												Reference Mark Names, with @Extended set to number
;				   +												of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefMarkDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefMarkList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMarks
	Local $asRefMarks[0]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oRefMarks = $oDoc.getReferenceMarks()
	If Not IsObj($oRefMarks) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$asRefMarks = $oRefMarks.getElementNames()
	If Not IsArray($asRefMarks) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	Return (UBound($asRefMarks) > 0) ? SetError($__LOW_STATUS_SUCCESS, UBound($asRefMarks), $asRefMarks) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldRefMarkList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefMarkSet
; Description ...: Create and Insert a Reference Mark at a Cursor position.
; Syntax ........: _LOWriter_FieldRefMarkSet(Byref $oDoc, Byref $oCursor, $sName[, $bOverwrite = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $sName               - a string value. The name of the Reference Mark to create.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $sName not a String.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = Document already contains a Reference Mark by the same name as called in
;				   +									$sName.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Reference Marks Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.text.ReferenceMark" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1. = Success. Successfully created a Reference Mark.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefMarkDelete, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefMarkSet(ByRef $oDoc, ByRef $oCursor, $sName, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMark, $oRefMarks

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$oRefMarks = $oDoc.getReferenceMarks()
	If Not IsObj($oRefMarks) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $oRefMarks.hasByName($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$oRefMark = $oDoc.createInstance("com.sun.star.text.ReferenceMark")
	If Not IsObj($oRefMark) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$oRefMark.Name = $sName

	$oCursor.Text.insertTextContent($oCursor, $oRefMark, $bOverwrite)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldRefMarkSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefModify
; Description ...: Set or Retrieve a Reference Field's settings.
; Syntax ........: _LOWriter_FieldRefModify(Byref $oDoc, Byref $oRefField[, $sRefMarkName = Null[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oRefField           - [in/out] an object. A Reference field Object from a previous Insert or retrieval
;				   +									function.
;                  $sRefMarkName        - [optional] a string value. Default is Null. The Reference Mark Name to Reference.
;                  $iRefUsing           - [optional] an integer value. Default is Null. The Type of reference to insert, see
;				   +									Constants.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oRefField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sRefMarkName not a String.
;				   @Error 1 @Extended 4 Return 0 = Document does not contain a Reference Mark by the same name as called in
;				   +									$sRefMarkName.
;				   @Error 1 @Extended 6 Return 0 = $iRefUsing not an Integer, Less than 0 or greater than 4. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Reference Marks Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $sRefMarkName
;				   |								2 = Error setting $iRefUsing
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;Refer Using Constants: $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED(0), The page number is displayed using Arabic numbers.
;							$LOW_FIELD_REF_USING_CHAPTER(1), The number of the chapter is displayed.
;							$LOW_FIELD_REF_USING_REF_TEXT(2), The reference text is displayed.
;							$LOW_FIELD_REF_USING_ABOVE_BELOW(3), The reference is displayed as one of the words, "above" or
;								"below".
;							$LOW_FIELD_REF_USING_PAGE_NUM_STYLED(4), The page number is displayed using the numbering type
;								defined in the page style of the reference position.
; Related .......: _LOWriter_FieldRefInsert, _LOWriter_FieldsGetList, _LOWriter_FieldRefMarkSet, _LOWriter_FieldRefMarkList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefModify(ByRef $oDoc, ByRef $oRefField, $sRefMarkName = Null, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMarks
	Local $iError = 0
	Local $avRef[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRefField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sRefMarkName, $iRefUsing) Then
		__LOWriter_ArrayFill($avRef, $oRefField.SourceName(), $oRefField.ReferenceFieldPart())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avRef)
	EndIf

	If ($sRefMarkName <> Null) Then
		If Not IsString($sRefMarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oRefMarks = $oDoc.getReferenceMarks()
		If Not IsObj($oRefMarks) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
		If Not $oRefMarks.hasByName($sRefMarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oRefField.SourceName = $sRefMarkName
		$oRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_REF_MARK ;Set Type to RefMark in case input field Obj is a diff type.
		$iError = ($oRefField.SourceName = $sRefMarkName) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oRefField.ReferenceFieldPart = $iRefUsing
		$iError = ($oRefField.ReferenceFieldPart = $iRefUsing) ? $iError : BitOR($iError, 2)
	EndIf

	$oRefField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldRefModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldsAdvGetList
; Description ...: Retrieve an Array of Advanced Field Objects contained in a document.
; Syntax ........: _LOWriter_FieldsAdvGetList(Byref $oDoc[, $iType = $LOW_FIELDADV_TYPE_ALL[, $bSupportedServices = True[, $bFieldType = True[, $bFieldTypeNum = True]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $iType               - [optional] an integer value. Default is $LOW_FIELDADV_TYPE_ALL. The type of Field to
;				   +						search for. See Constants. Can be BitOr'd together.
;                  $bSupportedServices  - [optional] a boolean value. Default is True. If True, adds a column to the array that
;				   +						has the supported service String for that particular Field, To assist in identifying
;				   +						the Field type.
;                  $bFieldType          - [optional] a boolean value. Default is True. If True, adds a column to the array that
;				   +						has the Field Type String for that particular Field as described by Libre Office.
;				   +						To assist in identifying the Field type.
;                  $bFieldTypeNum       - [optional] a boolean value. Default is True. If True, adds a column to the array that
;				   +						has the Field Type Constant Integer for that particular Field, to assist in
;				   +						identifying the Field type.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 1 or greater than 1023. (The total of
;				   +									all Constants added together.) See Constants.
;				   @Error 1 @Extended 3 Return 0 = $bSupportedServices not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bFieldType not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bFieldTypeNum not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $avFieldTypes passed to internal function not an Array. UDF needs fixed.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error converting Field type Constants.
;				   @Error 2 @Extended 2 Return 0 = Failed to create enumeration of fields in document.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning Array of Text Field Objects with @Extended set to
;				   +										number of results. See Remarks for Array sizing.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:The Array can vary in the number of columns, if $bSupportedServices, $bFieldType, and $bFieldTypeNum are set
;					to False, the Array will be a single column. With each of the above listed options being set to True, a
;					column will be added in the order they are listed in the UDF parameters. The First column will always be the
;					Field Object.
;					Setting $bSupportedServices to True will add a Supported Service String column for the found Field.
;					Setting $bFieldType to True will add a Field type column for the found Field.
;					Setting $bFieldTypeNum to True will add a Field type Number column, matching the below constants, for the
;						found Field.
;					Note: For simplicity, and also due to certain Bit limitations I have broken the different Field types into
;						three different categories, Regular Fields, ($LWFieldType), Advanced(Complex) Fields, ($LWFieldAdvType),
;						and Document Information fields (Found in the Document Information Tab in L.O. Fields dialog),
;						($LWFieldDocInfoType). Just because a field is listed in the constants below, does not necessarily mean
;						that I have made a function to create/modify it, you may still be able to update or delete it using the
;						Field Update, or Field Delete function, though. Some Fields are too complex to create a function for,
;						and others are literally impossible.
;Advanced Field Type Constants: $LOW_FIELDADV_TYPE_ALL = 1, All of the below listed Fields will be returned.
;								$LOW_FIELDADV_TYPE_BIBLIOGRAPHY, A Bibliography Field, found in Fields dialog, Database tab.
;								$LOW_FIELDADV_TYPE_DATABASE, A Database Field, found in Fields dialog, Database tab.
;								$LOW_FIELDADV_TYPE_DATABASE_SET_NUM, A Database Field, found in Fields dialog, Database tab.
;								$LOW_FIELDADV_TYPE_DATABASE_NAME, A Database Field, found in Fields dialog, Database tab.
;								$LOW_FIELDADV_TYPE_DATABASE_NEXT_SET, A Database Field, found in Fields dialog, Database tab.
;								$LOW_FIELDADV_TYPE_DATABASE_NAME_OF_SET, A Database Field, found in Fields dialog, Database tab.
;								$LOW_FIELDADV_TYPE_DDE, A DDE Field, found in Fields dialog, Variables tab.
;								$LOW_FIELDADV_TYPE_INPUT_USER, ?
;								$LOW_FIELDADV_TYPE_USER, A User Field, found in Fields dialog, Variables tab.
; Related .......: _LOWriter_FieldsDocInfoGetList, _LOWriter_FieldsGetList, _LOWriter_FieldDelete, _LOWriter_FieldGetAnchor,
;					_LOWriter_FieldUpdate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldsAdvGetList(ByRef $oDoc, $iType = $LOW_FIELDADV_TYPE_ALL, $bSupportedServices = True, $bFieldType = True, $bFieldTypeNum = True)
	Local $avFieldTypes[0][0]
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_IntIsBetween($iType, $LOW_FIELDADV_TYPE_ALL, 1023) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	;1023 is all possible Consts added together
	If Not IsBool($bSupportedServices) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bFieldType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bFieldTypeNum) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$avFieldTypes = __LOWriter_FieldTypeServices($iType, True, False)
	If @error > 0 Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$vReturn = __LOWriter_FieldsGetList($oDoc, $bSupportedServices, $bFieldType, $bFieldTypeNum, $avFieldTypes)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FieldsAdvGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldsDocInfoGetList
; Description ...: Retrieve an Array of Document Information Field Objects contained in a document.
; Syntax ........: _LOWriter_FieldsDocInfoGetList(Byref $oDoc[, $iType = $LOW_FIELD_DOCINFO_TYPE_ALL[, $bSupportedServices = True[, $bFieldType = True[, $bFieldTypeNum = True]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $iType               - [optional] an integer value. Default is $LOW_FIELD_DOCINFO_TYPE_ALL. The type of Field
;				   +						to search for. See Constants. Can be BitOr'd together.
;                  $bSupportedServices  - [optional] a boolean value. Default is True. If True, adds a column to the array that
;				   +						has the supported service String for that particular Field, To assist in identifying
;				   +						the Field type.
;                  $bFieldType          - [optional] a boolean value. Default is True. If True, adds a column to the array that
;				   +						has the Field Type String for that particular Field as described by Libre Office.
;				   +						To assist in identifying the Field type.
;                  $bFieldTypeNum       - [optional] a boolean value. Default is True. If True, adds a column to the array that
;				   +						has the Field Type Constant Integer for that particular Field, to assist in
;				   +						identifying the Field type.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 1 or greater than 16383. (The total of
;				   +									all Constants added together.) See Constants.
;				   @Error 1 @Extended 3 Return 0 = $bSupportedServices not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bFieldType not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bFieldTypeNum not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $avFieldTypes passed to internal function not an Array. UDF needs fixed.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error converting Field type Constants.
;				   @Error 2 @Extended 2 Return 0 = Failed to create enumeration of fields in document.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning Array of Text Field Objects with @Extended set to
;				   +										number of results. See Remarks for Array sizing.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:The Array can vary in the number of columns, if $bSupportedServices, $bFieldType, and $bFieldTypeNum are set
;					to False, the Array will be a single column. With each of the above listed options being set to True, a
;					column will be added in the order they are listed in the UDF parameters. The First column will always be the
;					Field Object.
;					Setting $bSupportedServices to True will add a Supported Service String column for the found Field.
;					Setting $bFieldType to True will add a Field type column for the found Field.
;					Setting $bFieldTypeNum to True will add a Field type Number column, matching the below constants, for the
;						found Field.
;					Note: For simplicity, and also due to certain Bit limitations I have broken the different Field types into
;						three different categories, Regular Fields, ($LWFieldType), Advanced(Complex) Fields, ($LWFieldAdvType),
;						and Document Information fields (Found in the Document Information Tab in L.O. Fields dialog),
;						($LWFieldDocInfoType). Just because a field is listed in the constants below, does not necessarily mean
;						that I have made a function to create/modify it, you may still be able to update or delete it using the
;						Field Update, or Field Delete function, though. Some Fields are too complex to create a function for,
;						and others are literally impossible.
;Doc Info Field Type Constants: $LOW_FIELD_DOCINFO_TYPE_ALL = 1, Returns a list of all field types listed below.
;									$LOW_FIELD_DOCINFO_TYPE_MOD_AUTH, A Modified By Author Field, found in Fields dialog,
;																	DocInformation Tab, Modified Type.
;									$LOW_FIELD_DOCINFO_TYPE_MOD_DATE_TIME,  A Modified Date/Time Field, found in Fields dialog,
;																	DocInformation Tab, Modified Type.
;									$LOW_FIELD_DOCINFO_TYPE_CREATE_AUTH, A Created By Author Field, found in Fields dialog,
;																	DocInformation Tab, Created Type.
;									$LOW_FIELD_DOCINFO_TYPE_CREATE_DATE_TIME, A Created Date/Time Field, found in Fields dialog,
;																	DocInformation Tab, Created Type.
;									$LOW_FIELD_DOCINFO_TYPE_CUSTOM, A Custom Field, found in Fields dialog, DocInformation Tab.
;									$LOW_FIELD_DOCINFO_TYPE_COMMENTS, A Comments Field, found in Fields dialog, DocInformation
;																	Tab.
;									$LOW_FIELD_DOCINFO_TYPE_EDIT_TIME, A Total Editing Time Field, found in Fields dialog,
;																	DocInformation Tab.
;									$LOW_FIELD_DOCINFO_TYPE_KEYWORDS, A Keywords Field, found in Fields dialog, DocInformation
;																	Tab.
;									$LOW_FIELD_DOCINFO_TYPE_PRINT_AUTH, A Printed By Author Field, found in Fields dialog,
;																	DocInformation Tab, Last Printed Type.
;									$LOW_FIELD_DOCINFO_TYPE_PRINT_DATE_TIME,  A Printed Date/Time Field, found in Fields dialog,
;																	DocInformation Tab, Last Printed Type.
;									$LOW_FIELD_DOCINFO_TYPE_REVISION, A Revision Number Field, found in Fields dialog,
;																	DocInformation Tab.
;									$LOW_FIELD_DOCINFO_TYPE_SUBJECT, A Subject Field, found in Fields dialog, DocInformation
;																	Tab.
;									$LOW_FIELD_DOCINFO_TYPE_TITLE, A Title Field, found in Fields dialog, DocInformation Tab.
; Related .......: _LOWriter_FieldsAdvGetList, _LOWriter_FieldsGetList, _LOWriter_FieldDelete, _LOWriter_FieldGetAnchor,
;					_LOWriter_FieldUpdate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldsDocInfoGetList(ByRef $oDoc, $iType = $LOW_FIELD_DOCINFO_TYPE_ALL, $bSupportedServices = True, $bFieldType = True, $bFieldTypeNum = True)
	Local $avFieldTypes[0][0]
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_IntIsBetween($iType, $LOW_FIELDADV_TYPE_ALL, 16383) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	;16383 is all possible Consts added together
	If Not IsBool($bSupportedServices) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bFieldType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bFieldTypeNum) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$avFieldTypes = __LOWriter_FieldTypeServices($iType, False, True)
	If @error > 0 Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$vReturn = __LOWriter_FieldsGetList($oDoc, $bSupportedServices, $bFieldType, $bFieldTypeNum, $avFieldTypes)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FieldsDocInfoGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSenderInsert
; Description ...: Insert a Sender Field.
; Syntax ........: _LOWriter_FieldSenderInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sContent = Null[, $iDataType = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sContent            - [optional] a string value. Default is Null. The Content to Display, only valid if
;				   +									$bIsFixed is set to True.
;                  $iDataType           - [optional] an integer value. Default is Null. The Data Type to display. See Constants.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sContent not a String.
;				   @Error 1 @Extended 7 Return 0 = $iDataType not an Integer, less than 0 or greater than 14. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.ExtendedUser" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Sender field, returning
;				   +										Sender Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Sender Data Type Constants: $LOW_FIELD_USER_DATA_COMPANY(0), The field shows the company name.
;								$LOW_FIELD_USER_DATA_FIRST_NAME(1), The field shows the first name.
;								$LOW_FIELD_USER_DATA_NAME(2), The field shows the name.
;								$LOW_FIELD_USER_DATA_SHORTCUT(3), The field shows the initials.
;								$LOW_FIELD_USER_DATA_STREET(4), The field shows the street.
;								$LOW_FIELD_USER_DATA_COUNTRY(5), The field shows the country.
;								$LOW_FIELD_USER_DATA_ZIP(6), The field shows the zip code.
;								$LOW_FIELD_USER_DATA_CITY(7), The field shows the city.
;								$LOW_FIELD_USER_DATA_TITLE(8), The field shows the title.
;								$LOW_FIELD_USER_DATA_POSITION(9), The field shows the position.
;								$LOW_FIELD_USER_DATA_PHONE_PRIVATE(10), The field shows the number of the private phone.
;								$LOW_FIELD_USER_DATA_PHONE_COMPANY(11), The field shows the number of the business phone.
;								$LOW_FIELD_USER_DATA_FAX(12), The field shows the fax number.
;								$LOW_FIELD_USER_DATA_EMAIL(13), The field shows the e-Mail.
;								$LOW_FIELD_USER_DATA_STATE(14), The field shows the state.
; Related .......: _LOWriter_FieldSenderModify,  _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSenderInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sContent = Null, $iDataType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSenderField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oSenderField = $oDoc.createInstance("com.sun.star.text.TextField.ExtendedUser")
	If Not IsObj($oSenderField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oSenderField.IsFixed = $bIsFixed
	EndIf

	If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oSenderField.Content = $sContent
	EndIf

	If ($iDataType <> Null) Then
		If Not __LOWriter_IntIsBetween($iDataType, $LOW_FIELD_USER_DATA_COMPANY, $LOW_FIELD_USER_DATA_STATE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oSenderField.UserDataType = $iDataType
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oSenderField, $bOverwrite)

	If ($sContent <> Null) Then ;Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oSenderField.Content <> $sContent And ($oSenderField.IsFixed() = True) Then $oSenderField.Content = $sContent
	EndIf

	If ($oSenderField.IsFixed() = False) Then $oSenderField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oSenderField)
EndFunc   ;==>_LOWriter_FieldSenderInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSenderModify
; Description ...: Set or Retrieve a Sender Field's settings.
; Syntax ........: _LOWriter_FieldSenderModify(Byref $oSenderField[, $bIsFixed = Null[, $sContent = Null[, $iDataType = Null]]])
; Parameters ....: $oSenderField        - [in/out] an object. A Sender field Object from a previous Insert or retrieval
;				   +									function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, this is the
;				   +								value does not update if the source changes or all fields are updated.
;                  $sContent            - [optional] a string value. Default is Null. The Content to Display, only valid if
;				   +									$bIsFixed is set to True.
;                  $iDataType           - [optional] an integer value. Default is Null. The Data Type to display. See Constants.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSenderField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sContent not a String.
;				   @Error 1 @Extended 4 Return 0 = $iDataType not an Integer, less than 0 or greater than 14. See Constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sContent
;				   |								4 = Error setting $iDataType
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;Sender Data Type Constants:  $LOW_FIELD_USER_DATA_COMPANY(0), The field shows the company name.
;								$LOW_FIELD_USER_DATA_FIRST_NAME(1), The field shows the first name.
;								$LOW_FIELD_USER_DATA_NAME(2), The field shows the name.
;								$LOW_FIELD_USER_DATA_SHORTCUT(3), The field shows the initials.
;								$LOW_FIELD_USER_DATA_STREET(4), The field shows the street.
;								$LOW_FIELD_USER_DATA_COUNTRY(5), The field shows the country.
;								$LOW_FIELD_USER_DATA_ZIP(6), The field shows the zip code.
;								$LOW_FIELD_USER_DATA_CITY(7), The field shows the city.
;								$LOW_FIELD_USER_DATA_TITLE(8), The field shows the title.
;								$LOW_FIELD_USER_DATA_POSITION(9), The field shows the position.
;								$LOW_FIELD_USER_DATA_PHONE_PRIVATE(10), The field shows the number of the private phone.
;								$LOW_FIELD_USER_DATA_PHONE_COMPANY(11), The field shows the number of the business phone.
;								$LOW_FIELD_USER_DATA_FAX(12), The field shows the fax number.
;								$LOW_FIELD_USER_DATA_EMAIL(13), The field shows the e-Mail.
;								$LOW_FIELD_USER_DATA_STATE(14), The field shows the state.
; Related .......: _LOWriter_FieldSenderInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSenderModify(ByRef $oSenderField, $bIsFixed = Null, $sContent = Null, $iDataType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avExtUser[3]

	If Not IsObj($oSenderField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sContent, $iDataType) Then
		__LOWriter_ArrayFill($avExtUser, $oSenderField.IsFixed(), $oSenderField.Content(), $oSenderField.UserDataType())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avExtUser)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oSenderField.IsFixed = $bIsFixed
		$iError = ($oSenderField.IsFixed() = $bIsFixed) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oSenderField.Content = $sContent
		$iError = ($oSenderField.Content() = $sContent) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iDataType <> Null) Then
		If Not __LOWriter_IntIsBetween($iDataType, $LOW_FIELD_USER_DATA_COMPANY, $LOW_FIELD_USER_DATA_STATE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oSenderField.UserDataType = $iDataType
		$iError = ($oSenderField.UserDataType() = $iDataType) ? $iError : BitOR($iError, 4)
	EndIf

	If ($oSenderField.IsFixed() = False) Then $oSenderField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldSenderModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarInsert
; Description ...: Insert a Set Variable Field.
; Syntax ........: _LOWriter_FieldSetVarInsert(Byref $oDoc, Byref $oCursor, $sName, $sValue[, $bOverwrite = False[, $iNumFormatKey = Null[, $bIsVisible = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $sName               - a string value. The name of the Set Variable Field to Create, If the name matches an
;				   +						already existing Set Variable Master Field, that Master will be used, else a new
;				   +						Set Variable Masterfield will be created.
;                  $sValue              - a string value. The Set Variable Field's value.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iNumFormatKey          - [optional] an integer value. Default is Null. The Number Format Key to use for
;				   +									displaying this variable.
;                  $bIsVisible          - [optional] a boolean value. Default is Null. If False, the Set Variable Field is
;				   +									invisible. L.O.'s default is True.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $sName not a String.
;				   @Error 1 @Extended 5 Return 0 = $sValue not a String.
;				   @Error 1 @Extended 6 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iNumFormatKeyKey not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iNumFormatKeyKey not equal to -1 and Number Format key called in
;				   +									$iNumFormatKeyKey not found in document.
;				   @Error 1 @Extended 9 Return 0 = $bIsVisible not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.SetExpression" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Set Variable field, returning
;				   +										Set Variable Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldSetVarModify,  _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_FormatKeyCreate _LOWriter_FormatKeyList, _LOWriter_FieldSetVarMasterCreate,
;					_LOWriter_FieldSetVarMasterList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarInsert(ByRef $oDoc, ByRef $oCursor, $sName, $sValue, $bOverwrite = False, $iNumFormatKey = Null, $bIsVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSetVarField, $oSetVarMaster
	Local $iExtended = 0

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sValue) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$oSetVarField = $oDoc.createInstance("com.sun.star.text.TextField.SetExpression")
	If Not IsObj($oSetVarField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If _LOWriter_FieldSetVarMasterExists($oDoc, $sName) Then
		$oSetVarMaster = _LOWriter_FieldSetVarMasterGetObj($oDoc, $sName)
		$iExtended = 1 ;1 = Master already existed.
	Else
		$oSetVarMaster = _LOWriter_FieldSetVarMasterCreate($oDoc, $sName)
	EndIf

	If Not IsObj($oSetVarMaster) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$oSetVarField.Content = $sValue

	If ($iNumFormatKey <> Null) Then
		If Not IsInt($iNumFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		If ($iNumFormatKey <> -1) And Not _LOWriter_FormatKeyExists($oDoc, $iNumFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oSetVarField.NumberFormat = $iNumFormatKey
	Else
		$oSetVarField.NumberFormat = 0 ; If No Input, set to General
	EndIf

	If ($bIsVisible <> Null) Then
		If Not IsBool($bIsVisible) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$oSetVarField.IsVisible = $bIsVisible
	EndIf

	$oSetVarField.attachTextFieldMaster($oSetVarMaster)

	$oCursor.Text.insertTextContent($oCursor, $oSetVarField, $bOverwrite)

	$oSetVarField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, $iExtended, $oSetVarField)
EndFunc   ;==>_LOWriter_FieldSetVarInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterCreate
; Description ...: Create a Set Variable Master Field.
; Syntax ........: _LOWriter_FieldSetVarMasterCreate(Byref $oDoc, $sMasterFieldName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sMasterFieldName    - a string value. The Set Variable Master Field name to create.
; Return values .:  Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sMasterFieldName not a String.
;				   @Error 1 @Extended 3 Return 0 = Document already contains a MasterField by the name called in
;				   +									$sMasterFieldName
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve MasterFields Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to Create MasterField Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully created the MasterField, returning MasterField
;				   +												Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldSetVarMasterDelete, _LOWriter_FieldSetVarInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterCreate(ByRef $oDoc, $sMasterFieldName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMasterFields, $oMasterfield
	Local $sFullFieldName
	Local $sField = "com.sun.star.text.fieldmaster.SetExpression"

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sMasterFieldName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$sFullFieldName = $sField & "." & $sMasterFieldName
	$oMasterFields = $oDoc.getTextFieldMasters()
	If Not IsObj($oMasterFields) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	If $oMasterFields.hasByName($sFullFieldName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$oMasterfield = $oDoc.createInstance($sField)
	If Not IsObj($oMasterfield) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$oMasterfield.Name = $sMasterFieldName

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oMasterfield)
EndFunc   ;==>_LOWriter_FieldSetVarMasterCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterDelete
; Description ...: Delete a Set Variable Master Field.
; Syntax ........: _LOWriter_FieldSetVarMasterDelete(Byref $oDoc, $vMasterField)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $vMasterField        - a variant value. The Set Variable Master Field name to delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $vMasterField not a String and not an Object.
;				   @Error 1 @Extended 3 Return 0 = $vMasterField is a String, but document does not contain a Masterfield by
;				   +									that name.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve MasterFields Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve MasterField object called in $vMasterField.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to delete MasterField, but document still contains a MasterField
;				   +									by that name.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested MasterField.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldSetVarMasterCreate, _LOWriter_FieldSetVarMasterGetObj, _LOWriter_FieldSetVarMasterList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterDelete(ByRef $oDoc, $vMasterField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMasterFields, $oMasterfield
	Local $sFullFieldName
	Local $sField = "com.sun.star.text.fieldmaster.SetExpression"

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($vMasterField) And Not IsObj($vMasterField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oMasterFields = $oDoc.getTextFieldMasters()
	If Not IsObj($oMasterFields) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If IsObj($vMasterField) Then
		$sFullFieldName = $sField & "." & $vMasterField.Name()
		$oMasterfield = $vMasterField
	Else
		$sFullFieldName = $sField & "." & $vMasterField
		If Not $oMasterFields.hasByName($sFullFieldName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oMasterfield = $oMasterFields.getByName($sFullFieldName)
	EndIf

	If Not IsObj($oMasterfield) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$oMasterfield.dispose()

	Return ($oMasterFields.hasByName($sFullFieldName)) ? SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldSetVarMasterDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterExists
; Description ...: Check if a document contains a Set Variable Master Field by name.
; Syntax ........: _LOWriter_FieldSetVarMasterExists(Byref $oDoc, $sMasterFieldName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sMasterFieldName    - a string value. The Set Variable Master Field name to look for.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sMasterFieldName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve MasterFields Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. If the document contains a MasterField by the called name,
;				   +												then True is returned, Else false.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterExists(ByRef $oDoc, $sMasterFieldName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMasterFields
	Local $sFullFieldName = "com.sun.star.text.fieldmaster.SetExpression" & "."

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sMasterFieldName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$sFullFieldName &= $sMasterFieldName

	$oMasterFields = $oDoc.getTextFieldMasters()
	If Not IsObj($oMasterFields) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	If $oMasterFields.hasByName($sFullFieldName) Then Return SetError($__LOW_STATUS_SUCCESS, 1, True)

	Return SetError($__LOW_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOWriter_FieldSetVarMasterExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterGetObj
; Description ...: Retrieve a Set Variable Master Field Object.
; Syntax ........: _LOWriter_FieldSetVarMasterGetObj(Byref $oDoc, $sMasterFieldName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sMasterFieldName    - a string value. The Set Variable Master Field to retrieve the Object for.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sMasterFieldName not an Object.
;				   @Error 1 @Extended 3 Return 0 = Document does not contain FieldMaster named as called in $sMasterFieldName.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve MasterFields Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve requested FieldMaster Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully retrieved requested FieldMaster Object. Returning
;				   +												requested Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldSetVarMasterList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterGetObj(ByRef $oDoc, $sMasterFieldName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMasterFields, $oMasterfield
	Local $sFullFieldName = "com.sun.star.text.fieldmaster.SetExpression" & "."

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sMasterFieldName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$sFullFieldName &= $sMasterFieldName

	$oMasterFields = $oDoc.getTextFieldMasters()
	If Not IsObj($oMasterFields) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	If Not $oMasterFields.hasByName($sFullFieldName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$oMasterfield = $oMasterFields.getByName($sFullFieldName)
	If Not IsObj($oMasterfield) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oMasterfield)
EndFunc   ;==>_LOWriter_FieldSetVarMasterGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterList
; Description ...: Retrieve a List of current Set Variable Master Fields in a document.
; Syntax ........: _LOWriter_FieldSetVarMasterList(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve MasterFields Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Array of MasterField Objects.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Successfully retrieved Array of Set Variable MasterField Names,
;				   +												returning Array of Set Variable MasterField Names with
;				   +												@Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:Note: This function includes in the list about 5 built-in Master Fields from Libre Office, namely:
;					Illustration, Table, Text, Drawing, and Figure.
; Related .......: _LOWriter_FieldSetVarMasterGetObj, _LOWriter_FieldSetVarMasterDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMasterFields
	Local $asMasterFields[0], $asSetVarMasters[0]
	Local $iCount = 0
	Local $sField = "com.sun.star.text.fieldmaster.SetExpression"

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oMasterFields = $oDoc.getTextFieldMasters()
	If Not IsObj($oMasterFields) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$asMasterFields = $oMasterFields.getElementNames()
	If Not IsArray($asMasterFields) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	ReDim $asSetVarMasters[UBound($asMasterFields)]

	For $i = 0 To UBound($asMasterFields) - 1
		If ($oMasterFields.getByName($asMasterFields[$i]).supportsService($sField)) Then
			$asSetVarMasters[$iCount] = $oMasterFields.getByName($asMasterFields[$i]).Name()
			$iCount += 1
		EndIf

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next

	ReDim $asSetVarMasters[$iCount]

	Return SetError($__LOW_STATUS_SUCCESS, $iCount, $asSetVarMasters)
EndFunc   ;==>_LOWriter_FieldSetVarMasterList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterListFields
; Description ...: Return an Array of Objects of dependent fields for a specific Master Field.
; Syntax ........: _LOWriter_FieldSetVarMasterListFields(Byref $oDoc, Byref $oMasterfield)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oMasterfield        - [in/out] an object. The Set Variable Master Field Object returned from a previous
;				   +						Create or retrieval function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oMasterfield not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve dependent fields Array.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully searched for dependent fields, but MasterField does
;				   +											not contain any.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for dependent fields, returning Array of
;				   +												dependent SetVariable Fields, with @Extended set to number
;				   +												of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Dependent Fields are SetVariable Fields that are referencing the Master field.
; Related .......: _LOWriter_FieldSetVarMasterGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterListFields(ByRef $oDoc, ByRef $oMasterfield)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aoDependFields[0]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oMasterfield) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$aoDependFields = $oMasterfield.DependentTextFields()
	If Not IsArray($aoDependFields) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return (UBound($aoDependFields) > 0) ? SetError($__LOW_STATUS_SUCCESS, UBound($aoDependFields), $aoDependFields) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldSetVarMasterListFields

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarModify
; Description ...: Set or Retrieve a Set Variable Field's settings.
; Syntax ........: _LOWriter_FieldSetVarModify(Byref $oDoc, Byref $oSetVarField[, $sValue = Null[, $iNumFormatKey = Null[, $bIsVisible = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oSetVarField        - [in/out] an object. A Set Variable field Object from a previous Insert or retrieval
;				   +									function.
;                  $sValue              - [optional] a string value. Default is Null. The Set Variable Field's value.
;                  $iNumFormatKey       - [optional] an integer value. Default is Null. The Number Format Key to use for
;				   +									displaying this variable.
;                  $bIsVisible          - [optional] a boolean value. Default is Null. If False, the Set Variable Field is
;				   +									invisible. L.O.'s default is True.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSetVarField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sValue not a String.
;				   @Error 1 @Extended 4 Return 0 = $iNumFormatKeyKey not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iNumFormatKeyKey not equal to -1 and Number Format key called in
;				   +									$iNumFormatKeyKey not found in document.
;				   @Error 1 @Extended 6 Return 0 = $bIsVisible not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $sValue
;				   |								2 = Error setting $iNumFormatKey
;				   |								4 = Error setting $bIsVisible
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
;				   +								The fourth element is the Variable Name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldSetVarInsert, _LOWriter_FieldsGetList, _LOWriter_FormatKeyCreate _LOWriter_FormatKeyList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarModify(ByRef $oDoc, ByRef $oSetVarField, $sValue = Null, $iNumFormatKey = Null, $bIsVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avSetVar[4]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSetVarField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sValue, $iNumFormatKey, $bIsVisible) Then
		;Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		;fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		;from the value.
		$iNumberFormat = $oSetVarField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : $iNumberFormat

		__LOWriter_ArrayFill($avSetVar, $oSetVarField.Content(), $iNumberFormat, $oSetVarField.IsVisible(), $oSetVarField.VariableName())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avSetVar)
	EndIf

	If ($sValue <> Null) Then
		If Not IsString($sValue) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oSetVarField.Content = $sValue
		$iError = ($oSetVarField.Content() = $sValue) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iNumFormatKey <> Null) Then
		If Not IsInt($iNumFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		If ($iNumFormatKey <> -1) And Not _LOWriter_FormatKeyExists($oDoc, $iNumFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oSetVarField.NumberFormat = $iNumFormatKey
		$iError = ($oSetVarField.NumberFormat() = $iNumFormatKey) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bIsVisible <> Null) Then
		If Not IsBool($bIsVisible) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oSetVarField.IsVisible = $bIsVisible
		$iError = ($oSetVarField.IsVisible() = $bIsVisible) ? $iError : BitOR($iError, 4)
	EndIf

	$oSetVarField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldSetVarModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldsGetList
; Description ...: Retrieve an Array of Field Objects contained in a document.
; Syntax ........: _LOWriter_FieldsGetList(Byref $oDoc[, $iType = $LOW_FIELD_TYPE_ALL[, $bSupportedServices = True[, $bFieldType = True[, $bFieldTypeNum = True]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $iType               - [optional] an integer value. Default is $LOW_FIELD_TYPE_ALL. The type of Field to
;				   +						search for. See Constants. Can be BitOr'd together.
;                  $bSupportedServices  - [optional] a boolean value. Default is True. If True, adds a column to the array that
;				   +						has the supported service String for that particular Field, To assist in identifying
;				   +						the Field type.
;                  $bFieldType          - [optional] a boolean value. Default is True. If True, adds a column to the array that
;				   +						has the Field Type String for that particular Field as described by Libre Office.
;				   +						To assist in identifying the Field type.
;                  $bFieldTypeNum       - [optional] a boolean value. Default is True. If True, adds a column to the array that
;				   +						has the Field Type Constant Integer for that particular Field, to assist in
;				   +						identifying the Field type.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 1 or greater than 2147483647. (The total of
;				   +									all Constants added together.) See Constants.
;				   @Error 1 @Extended 3 Return 0 = $bSupportedServices not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bFieldType not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bFieldTypeNum not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $avFieldTypes passed to internal function not an Array. UDF needs fixed.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error converting Field type Constants.
;				   @Error 2 @Extended 2 Return 0 = Failed to create enumeration of fields in document.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning Array of Text Field Objects with @Extended set to
;				   +										number of results. See Remarks for Array sizing.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array can vary in the number of columns, if $bSupportedServices, $bFieldType, and $bFieldTypeNum are set
;					to False, the Array will be a single column. With each of the above listed options being set to True, a
;					column will be added in the order they are listed in the UDF parameters. The First column will always be the
;					Field Object.
;					Setting $bSupportedServices to True will add a Supported Service String column for the found Field.
;					Setting $bFieldType to True will add a Field type column for the found Field.
;					Setting $bFieldTypeNum to True will add a Field type Number column, matching the below constants, for the
;						found Field.
;					Note: For simplicity, and also due to certain Bit limitations I have broken the different Field types into
;						three different categories, Regular Fields, ($LWFieldType), Advanced(Complex) Fields, ($LWFieldAdvType),
;						and Document Information fields (Found in the Document Information Tab in L.O. Fields dialog),
;						($LWFieldDocInfoType). Just because a field is listed in the constants below, does not necessarily mean
;						that I have made a function to create/modify it, you may still be able to update or delete it using the
;						Field Update, or Field Delete function, though. Some Fields are too complex to create a function for,
;						and others are literally impossible.
;Field Type Constants: $LOW_FIELD_TYPE_ALL = 1, Returns a list of all field types listed below.
;						$LOW_FIELD_TYPE_COMMENT, A Comment Field. As Found at Insert > Comment
;						$LOW_FIELD_TYPE_AUTHOR, A Author field, found in the Fields Dialog, Document tab.
;						$LOW_FIELD_TYPE_CHAPTER, A Chapter field, found in the Fields Dialog, Document tab.
;						$LOW_FIELD_TYPE_CHAR_COUNT, A Character Count field, found in the Fields Dialog, Document tab,
;							Statistics Type.
;						$LOW_FIELD_TYPE_COMBINED_CHAR, A Combined Character field, found in the Fields Dialog, Functions tab.
;						$LOW_FIELD_TYPE_COND_TEXT, A Conditional Text field, found in the Fields Dialog, Functions tab.
;						$LOW_FIELD_TYPE_DATE_TIME, A Date/Time field, found in the Fields Dialog, Document tab, Date Type and
;							Time Type.
;						$LOW_FIELD_TYPE_INPUT_LIST,  A Input List field, found in the Fields Dialog, Functions tab.
;						$LOW_FIELD_TYPE_EMB_OBJ_COUNT,  A Object Count field, found in the Fields Dialog, Document tab,
;							Statistics Type.
;						$LOW_FIELD_TYPE_SENDER, A Sender field, found in the Fields Dialog, Document tab.
;						$LOW_FIELD_TYPE_FILENAME,  A File Name field, found in the Fields Dialog, Document tab.
;						$LOW_FIELD_TYPE_SHOW_VAR, A Show Variable field, found in the Fields Dialog, Variables tab.
;						$LOW_FIELD_TYPE_INSERT_REF,  A Insert Reference field, found in the Fields Dialog, Cross-References tab,
;													including: "Insert Reference", "Headings", "Numbered Paragraphs", "Drawing",
;													"Bookmarks", "Footnotes", "Endnotes", etc.
;						$LOW_FIELD_TYPE_IMAGE_COUNT, A Image Count field, found in the Fields Dialog, Document tab, Statistics
;							Type.
;						$LOW_FIELD_TYPE_HIDDEN_PAR, A Hidden Paragraph field, found in the Fields Dialog, Functions tab.
;						$LOW_FIELD_TYPE_HIDDEN_TEXT, A Hidden Text field, found in the Fields Dialog, Functions tab.
;						$LOW_FIELD_TYPE_INPUT, A Input field, found in the Fields Dialog, Functions tab.
;						$LOW_FIELD_TYPE_PLACEHOLDER, A Placeholder field, found in the Fields Dialog, Functions tab.
;						$LOW_FIELD_TYPE_MACRO, A Execute Macro field, found in the Fields Dialog, Functions tab.
;						$LOW_FIELD_TYPE_PAGE_COUNT, A Page Count field, found in the Fields Dialog, Document tab, Statistics
;							Type.
;						$LOW_FIELD_TYPE_PAGE_NUM,  A Page Number (Unstyled) field, found in the Fields Dialog, Document tab.
;						$LOW_FIELD_TYPE_PAR_COUNT, A Paragraph Count field, found in the Fields Dialog, Document tab, Statistics
;							Type.
;						$LOW_FIELD_TYPE_SHOW_PAGE_VAR, A Show Page Variable field, found in the Fields Dialog, Variables tab.
;						$LOW_FIELD_TYPE_SET_PAGE_VAR, A Set Page Variable field, found in the Fields Dialog, Variables tab.
;						$LOW_FIELD_TYPE_SCRIPT,
;						$LOW_FIELD_TYPE_SET_VAR, A Set Variable field, found in the Fields Dialog, Variables tab.
;						$LOW_FIELD_TYPE_TABLE_COUNT, A Table Count field, found in the Fields Dialog, Document tab, Statistics
;							Type.
;						$LOW_FIELD_TYPE_TEMPLATE_NAME, A Templates field, found in the Fields Dialog, Document tab.
;						$LOW_FIELD_TYPE_URL,
;						$LOW_FIELD_TYPE_WORD_COUNT, A Word Count field, found in the Fields Dialog, Document tab, Statistics
;							Type.
; Related .......: _LOWriter_FieldsAdvGetList, _LOWriter_FieldsDocInfoGetList, _LOWriter_FieldDelete, _LOWriter_FieldGetAnchor,
;					_LOWriter_FieldUpdate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldsGetList(ByRef $oDoc, $iType = $LOW_FIELD_TYPE_ALL, $bSupportedServices = True, $bFieldType = True, $bFieldTypeNum = True)
	Local $avFieldTypes[0][0]
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_IntIsBetween($iType, $LOW_FIELD_TYPE_ALL, 2147483647) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	;2147483647 is all possible Consts added together
	If Not IsBool($bSupportedServices) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bFieldType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bFieldTypeNum) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$avFieldTypes = __LOWriter_FieldTypeServices($iType)
	If (@error > 0) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$vReturn = __LOWriter_FieldsGetList($oDoc, $bSupportedServices, $bFieldType, $bFieldTypeNum, $avFieldTypes)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FieldsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldShowVarInsert
; Description ...: Insert a Show Variable Field.
; Syntax ........: _LOWriter_FieldShowVarInsert(Byref $oDoc, Byref $oCursor, $sSetVarName[, $bOverwrite = False[, $iNumFormatKey = Null[, $bShowName = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $sSetVarName         - a string value. The Set Variable name to show the value of.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iNumFormatKey       - [optional] an integer value. Default is Null. The Number Format Key to display the
;				   +						content in
;                  $bShowName           - [optional] a boolean value. Default is Null. If True, the Set Variable name is
;				   +						displayed, rather than its value.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sSetVarName not a String.
;				   @Error 1 @Extended 6 Return 0 = Did not find a Set Var Field Master with same name as $sSetVarName.
;				   @Error 1 @Extended 7 Return 0 = $iNumFormatKey not an Integer.
;				   @Error 1 @Extended 8 Return 0 = Number Format key called in $iNumFormatKey not found in document.
;				   @Error 1 @Extended 9 Return 0 = $bShowName not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.GetExpression" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Show Variable field, returning
;				   +										Show Variable Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: This function checks if there is a Set Variable matching the name called in $sSetVarName.
; Related .......: _LOWriter_FieldShowVarModify, _LOWriter_FieldSetVarInsert, _LOWriter_FieldsGetList,
;					_LOWriter_FormatKeyCreate _LOWriter_FormatKeyList, _LOWriter_DocGetViewCursor,
;					_LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor,
;					 _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor,
;					_LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldShowVarInsert(ByRef $oDoc, ByRef $oCursor, $sSetVarName, $bOverwrite = False, $iNumFormatKey = Null, $bShowName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShowVarField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oShowVarField = $oDoc.createInstance("com.sun.star.text.TextField.GetExpression")
	If Not IsObj($oShowVarField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If Not IsString($sSetVarName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If Not _LOWriter_FieldSetVarMasterExists($oDoc, $sSetVarName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	$oShowVarField.Content = $sSetVarName

	If ($iNumFormatKey <> Null) Then
		If Not IsInt($iNumFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		If Not _LOWriter_FormatKeyExists($oDoc, $iNumFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oShowVarField.NumberFormat = $iNumFormatKey
	EndIf

	If ($bShowName <> Null) Then
		If Not IsBool($bShowName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$oShowVarField.IsShowFormula = $bShowName
		If ($bShowName = True) Then $oShowVarField.NumberFormat = -1
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oShowVarField, $bOverwrite)

	$oShowVarField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oShowVarField)
EndFunc   ;==>_LOWriter_FieldShowVarInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldShowVarModify
; Description ...:Set or Retrieve a Show Variable Field's settings.
; Syntax ........: _LOWriter_FieldShowVarModify(Byref $oDoc, Byref $oShowVarField[, $sSetVarName = Null[, $iNumFormatKey = Null[, $bShowName = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oShowVarField       - [in/out] an object. A Show Variable field Object from a previous Insert or retrieval
;				   +									function.
;                  $sSetVarName         - [optional] a string value. Default is Null. The Set Variable name to show the value of.
;                  $iNumFormatKey       - [optional] an integer value. Default is Null. The Number Format Key to display the
;				   +						content in
;                  $bShowName           - [optional] a boolean value. Default is Null. If True, the Set Variable name is
;				   +						displayed, rather than its value.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oShowVarField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sSetVarName not a String.
;				   @Error 1 @Extended 4 Return 0 = Did not find a Set Var Field Master with same name as $sSetVarName.
;				   @Error 1 @Extended 5 Return 0 = $iNumFormatKey not an Integer.
;				   @Error 1 @Extended 6 Return 0 = Number Format key called in $iNumFormatKey not found in document.
;				   @Error 1 @Extended 7 Return 0 = $bShowName not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4
;				   |								1 = Error setting $sSetVarName
;				   |								2 = Error setting $iNumFormatKey
;				   |								4 = Error setting $bShowName
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					 Note: This function checks if there is a Set Variable matching the name called in $sSetVarName.
; Related .......: _LOWriter_FieldShowVarInsert, _LOWriter_FieldsGetList, _LOWriter_FormatKeyCreate,  _LOWriter_FormatKeyList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldShowVarModify(ByRef $oDoc, ByRef $oShowVarField, $sSetVarName = Null, $iNumFormatKey = Null, $bShowName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avShowVar[3]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShowVarField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sSetVarName, $iNumFormatKey, $bShowName) Then
		;Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		;fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		;from the value.
		$iNumberFormat = $oShowVarField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : $iNumberFormat

		__LOWriter_ArrayFill($avShowVar, $oShowVarField.Content(), $iNumberFormat, $oShowVarField.IsShowFormula())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avShowVar)
	EndIf

	If ($sSetVarName <> Null) Then
		If Not IsString($sSetVarName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		If Not _LOWriter_FieldSetVarMasterExists($oDoc, $sSetVarName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oShowVarField.Content = $sSetVarName
		$iError = ($oShowVarField.Content() = $sSetVarName) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iNumFormatKey <> Null) Then
		If Not IsInt($iNumFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		If Not _LOWriter_FormatKeyExists($oDoc, $iNumFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oShowVarField.NumberFormat = $iNumFormatKey
		$iError = ($oShowVarField.NumberFormat() = ($iNumFormatKey)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bShowName <> Null) Then
		If Not IsBool($bShowName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oShowVarField.IsShowFormula = $bShowName
		$iError = ($oShowVarField.IsShowFormula() = $bShowName) ? $iError : BitOR($iError, 4)
	EndIf

	$oShowVarField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldShowVarModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldStatCountInsert
; Description ...: Insert a Count Field.
; Syntax ........: _LOWriter_FieldStatCountInsert(Byref $oDoc, Byref $oCursor, $iCountType[, $bOverwrite = False[, $iNumFormat = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $iCountType          - an integer value. The Type of Data to Count. See Constants.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iNumFormat            - [optional] an integer value. Default is Null. The numbering format to use for Count
;				   +						field numbering. See Constants.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $iCountType not an integer, less than 0 or greater than 6. See Constants.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iNumFormat not an integer, less than 0 or greater than 71. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create requested Count Field Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve Field Count Service Type. Check Constants.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Count Field. Returning
;				   +											the Count Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: After insertion there seems to be a necessary delay before the value to display is available, thus when a
;						new count field is inserted, the value will be "0". If you call a _LOWriter_FieldUpdate for this
;						field after a few seconds, the value should appear.
;Field Count Type Constants: $LOW_FIELD_COUNT_TYPE_CHARACTERS(0), Count field is a Character Count type field.
;								$LOW_FIELD_COUNT_TYPE_IMAGES(1), Count field is an Image Count type field.
;								$LOW_FIELD_COUNT_TYPE_OBJECTS(2),  Count field is an Object Count type field._
;								$LOW_FIELD_COUNT_TYPE_PAGES(3), Count field is a Page Count type field.
;								$LOW_FIELD_COUNT_TYPE_PARAGRAPHS(4), Count field is a Paragraph Count type field.
;								$LOW_FIELD_COUNT_TYPE_TABLES(5), Count field is a Table Count type field.
;								$LOW_FIELD_COUNT_TYPE_WORDS(6), Count field is a Word Count type field.
;Numbering Format Constants: $LOW_NUM_STYLE_CHARS_UPPER_LETTER(0), Numbering is put in upper case letters. ("A, B, C, D)
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
; Related .......: _LOWriter_FieldStatCountModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldStatCountInsert(ByRef $oDoc, ByRef $oCursor, $iCountType, $bOverwrite = False, $iNumFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCountField
	Local $sFieldType

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not __LOWriter_IntIsBetween($iCountType, $LOW_FIELD_COUNT_TYPE_CHARACTERS, $LOW_FIELD_COUNT_TYPE_WORDS) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$sFieldType = __LOWriter_FieldCountType($iCountType)
	If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	$oCountField = $oDoc.createInstance($sFieldType)
	If Not IsObj($oCountField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oCountField.NumberingType = $iNumFormat
	Else
		$oCountField.NumberingType = $LOW_NUM_STYLE_PAGE_DESCRIPTOR
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oCountField, $bOverwrite)

	$oCountField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oCountField)
EndFunc   ;==>_LOWriter_FieldStatCountInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldStatCountModify
; Description ...: Set or Retrieve a Count Field's settings.
; Syntax ........: _LOWriter_FieldStatCountModify(Byref $oDoc, Byref $oCountField[, $iCountType = Null[, $iNumFormat = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCountField         - [in/out] an object. A Count field Object from a previous Insert or retrieval
;				   +									function.
;                  $iCountType          - [optional] an integer value. Default is Null. The Type of Data to Count. See Constants.
;                  $iNumFormat            - [optional] an integer value. Default is Null. The numbering format to use for Count
;				   +						field numbering. See Constants.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCountField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iCountType not an integer, less than 0 or greater than 6. See Constants.
;				   @Error 1 @Extended 4 Return 0 = $iNumFormat not an integer, less than 0 or greater than 71. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create requested Count Field Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve Field Count Service Type. Check Constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $iCountType
;				   |								2 = Error setting $iNumFormat
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					 After changing the Count type there may be a delay before the value to display is available,
;						thus when the count field is inserted, the value will be "0". If you call a _LOWriter_FieldUpdate for
;						this field after a few seconds, the value should appear.
;Field Count Type Constants: $LOW_FIELD_COUNT_TYPE_CHARACTERS(0), Count field is a Character Count type field.
;								$LOW_FIELD_COUNT_TYPE_IMAGES(1), Count field is an Image Count type field.
;								$LOW_FIELD_COUNT_TYPE_OBJECTS(2),  Count field is an Object Count type field._
;								$LOW_FIELD_COUNT_TYPE_PAGES(3), Count field is a Page Count type field.
;								$LOW_FIELD_COUNT_TYPE_PARAGRAPHS(4), Count field is a Paragraph Count type field.
;								$LOW_FIELD_COUNT_TYPE_TABLES(5), Count field is a Table Count type field.
;								$LOW_FIELD_COUNT_TYPE_WORDS(6), Count field is a Word Count type field.
;Numbering Format Constants: $LOW_NUM_STYLE_CHARS_UPPER_LETTER(0), Numbering is put in upper case letters. ("A, B, C, D)
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
; Related .......: _LOWriter_FieldStatCountInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldStatCountModify(ByRef $oDoc, ByRef $oCountField, $iCountType = Null, $iNumFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avCountField[2]
	Local $oNewCountField
	Local $sFieldType

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCountField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($iNumFormat) Then
		__LOWriter_ArrayFill($avCountField, __LOWriter_FieldCountType($oCountField), $oCountField.NumberingType())
		If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avCountField)
	EndIf

	If ($iCountType <> Null) Then
		If Not __LOWriter_IntIsBetween($iCountType, $LOW_FIELD_COUNT_TYPE_CHARACTERS, $LOW_FIELD_COUNT_TYPE_WORDS) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$sFieldType = __LOWriter_FieldCountType($iCountType)
		If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

		If Not $oCountField.supportsService($sFieldType) Then ;If the Field is already that type, skip this and do nothing.

			$oNewCountField = $oDoc.createInstance($sFieldType)
			If Not IsObj($oNewCountField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

			;It doesn't work to just set a new Count type for an already inserted Count Field, so I have to create a new one and then
			;insert it.
			$oNewCountField.NumberingType = $oCountField.NumberingType()

			$oDoc.Text.createTextCursorByRange($oCountField.Anchor()).Text.insertTextContent($oCountField.Anchor(), $oNewCountField, True)

			;Update the Old Count Field Object to the new one.
			$oCountField = $oNewCountField

			$oCountField.Update()

			$iError = ($oCountField.supportsService($sFieldType)) ? $iError : BitOR($iError, 1)
		EndIf
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oCountField.NumberingType = $iNumFormat
		$iError = ($oCountField.NumberingType() = $iNumFormat) ? $iError : BitOR($iError, 2)
	EndIf

	$oCountField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldStatCountModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldStatTemplateInsert
; Description ...: Insert a Template Field.
; Syntax ........: _LOWriter_FieldStatTemplateInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $iFormat = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iFormat             - [optional] an integer value. Default is Null. The Format to display the Template data
;				   +									in. See Constants.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iFormat not an integer, less than 0 or greater than 5. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.TemplateName" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Template Field. Returning
;				   +											the Template Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;File Name Constants: $LOW_FIELD_FILENAME_FULL_PATH(0), The content of the URL is completely displayed.
;						$LOW_FIELD_FILENAME_PATH(1), Only the path of the file is displayed.
;						$LOW_FIELD_FILENAME_NAME(2), Only the name of the file without the file extension is displayed.
;						$LOW_FIELD_FILENAME_NAME_AND_EXT(3), The file name including the file extension is displayed.
;						$LOW_FIELD_FILENAME_CATEGORY(4), The Category of the Template is displayed.
;						$LOW_FIELD_FILENAME_TEMPLATE_NAME(5), The Template Name is displayed.
; Related .......: _LOWriter_FieldStatTemplateModify, _LOWriter_DocGenPropTemplate,  _LOWriter_DocGetViewCursor,
;					_LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor,
;					_LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor,
;					_LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldStatTemplateInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTemplateField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oTemplateField = $oDoc.createInstance("com.sun.star.text.TextField.TemplateName")
	If Not IsObj($oTemplateField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($iFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iFormat, $LOW_FIELD_FILENAME_FULL_PATH, $LOW_FIELD_FILENAME_TEMPLATE_NAME) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oTemplateField.FileFormat = $iFormat
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oTemplateField, $bOverwrite)

	$oTemplateField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oTemplateField)
EndFunc   ;==>_LOWriter_FieldStatTemplateInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldStatTemplateModify
; Description ...: Set or Retrieve a Template Field's settings.
; Syntax ........: _LOWriter_FieldStatTemplateModify(Byref $oTemplateField[, $iFormat = Null])
; Parameters ....: $oTemplateField      - [in/out] an object. A Template field Object from a previous Insert or retrieval
;				   +									function.
;                  $iFormat             - [optional] an integer value. Default is Null. The Format to display the Template data
;				   +									in. See Constants.
; Return values .: Success: 1 or Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oTemplateField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormat not an integer, less than 0 or greater than 5. See Constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1
;				   |								1 = Error setting $iFormat
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current
;				   +								Template Format Type setting, in Integer format. See File Name Constants.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;File Name Constants: $LOW_FIELD_FILENAME_FULL_PATH(0), The content of the URL is completely displayed.
;						$LOW_FIELD_FILENAME_PATH(1), Only the path of the file is displayed.
;						$LOW_FIELD_FILENAME_NAME(2), Only the name of the file without the file extension is displayed.
;						$LOW_FIELD_FILENAME_NAME_AND_EXT(3), The file name including the file extension is displayed.
;						$LOW_FIELD_FILENAME_CATEGORY(4), The Category of the Template is displayed.
;						$LOW_FIELD_FILENAME_TEMPLATE_NAME(5), The Template Name is displayed.
; Related .......: _LOWriter_FieldStatTemplateInsert, _LOWriter_DocGenPropTemplate, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldStatTemplateModify(ByRef $oTemplateField, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oTemplateField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iFormat) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oTemplateField.FileFormat())

	If Not __LOWriter_IntIsBetween($iFormat, $LOW_FIELD_FILENAME_FULL_PATH, $LOW_FIELD_FILENAME_TEMPLATE_NAME) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oTemplateField.FileFormat = $iFormat
	$iError = ($oTemplateField.FileFormat() = $iFormat) ? $iError : BitOR($iError, 1)

	$oTemplateField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldStatTemplateModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldUpdate
; Description ...: Update a Field or all fields in a document.
; Syntax ........: _LOWriter_FieldUpdate(Byref $oDoc[, $oField = Null[, $bForceUpdate = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oField              - [optional] an object. Default is Null. A Field Object returned from a previous Insert
;				   +						or Retrieve function. If left as Null, all Fields will be updated.
;                  $bForceUpdate        - [optional] a boolean value. Default is False. If True, Field(s) will be updated whether
;				   +						it(they) is(are) fixed or not. See remarks.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object
;				   @Error 1 @Extended 2 Return 0 = $oField not set to Null, and not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bForceUpdate not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve enumeration of all fields.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully updated requested field.
;				   @Error 0 @Extended 1 Return 1 = Success. Requested field is set to Fixed and $bForceUpdate is set to false,
;				   +									Field was not updated.
;				   @Error 0 @Extended ? Return 1 = Success. Successfully updated all fields, @Extended set to number of fields
;				   +											updated.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Updating a fixed field will usually erase any user-provided content, such as an author name, creation date
;						etc. If a Field is fixed, the field wont be updated unless $bForceUpdate is set to true.
; Related .......: _LOWriter_FieldsGetList _LOWriter_FieldsAdvGetList _LOWriter_FieldsDocInfoGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldUpdate(ByRef $oDoc, $oField = Null, $bForceUpdate = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextFields, $oTextField
	Local $iCount = 0, $iUpdated = 0

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If $oField <> Null And Not IsObj($oField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bForceUpdate) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If ($oField <> Null) Then
		If ($oField.getPropertySetInfo.hasPropertyByName("IsFixed") = True) Then
			If ($oField.IsFixed() = True) And ($bForceUpdate = False) Then Return SetError($__LOW_STATUS_SUCCESS, 1, 1) ;Updating a fixed field, causes its content to be removed.
		EndIf
		$oField.Update()
		Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
	EndIf

	;Update All Fields.
	$oTextFields = $oDoc.getTextFields.createEnumeration()
	If Not IsObj($oTextFields) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	While $oTextFields.hasMoreElements()
		$oTextField = $oTextFields.nextElement()

		If ($bForceUpdate = False) Then
			If ($oTextField.getPropertySetInfo.hasPropertyByName("IsFixed") = True) Then
				If ($oTextField.IsFixed() = False) Then
					$oTextField.Update()
					$iUpdated += 1
				EndIf ;Updating a fixed field, causes its content to be removed.
			Else
				$oTextField.Update()
				$iUpdated += 1
			EndIf


		Else
			$oTextField.Update()
			$iUpdated += 1
		EndIf

		$iCount += 1
		Sleep((IsInt($iCount / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	WEnd

	Return SetError($__LOW_STATUS_SUCCESS, $iUpdated, 1)
EndFunc   ;==>_LOWriter_FieldUpdate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldVarSetPageInsert
; Description ...: Insert a Set Page Variable Field.
; Syntax ........: _LOWriter_FieldVarSetPageInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $bRefOn = Null[, $iOffset = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $bRefOn              - [optional] a boolean value. Default is Null. If True, Reference point is enabled, else
;				   +									disabled.
;                  $iOffset             - [optional] an integer value. Default is Null. The offset the start the page count from.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bRefOn not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iOffset not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.ReferencePageSet" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Set Page Variable Field. Returning
;				   +											the Set Page Variable Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldVarSetPageModify, _LOWriter_DocGetViewCursor,	_LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldVarSetPageInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bRefOn = Null, $iOffset = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageVarSetField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oPageVarSetField = $oDoc.createInstance("com.sun.star.text.TextField.ReferencePageSet")
	If Not IsObj($oPageVarSetField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bRefOn <> Null) Then
		If Not IsBool($bRefOn) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oPageVarSetField.On = $bRefOn
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oPageVarSetField.Offset = $iOffset
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oPageVarSetField, $bOverwrite)

	$oPageVarSetField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oPageVarSetField)
EndFunc   ;==>_LOWriter_FieldVarSetPageInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldVarSetPageModify
; Description ...: Set or retrieve a Set Page Variable Field's settings.
; Syntax ........: _LOWriter_FieldVarSetPageModify(Byref $oPageVarSetField[, $bRefOn = Null[, $iOffset = Null]]])
; Parameters ....: $oPageVarSetField    - [in/out] an object. A Set Page Variable field Object from a previous Insert or
;				   +									retrieval function.
;                  $bRefOn              - [optional] a boolean value. Default is Null. If True, Reference point is enabled, else
;				   +									disabled.
;                  $iOffset             - [optional] an integer value. Default is Null. The offset the start the page count from.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oPageVarSetField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bRefOn not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $iOffset not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $bRefOn
;				   |								2 = Error setting $iOffset
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldVarSetPageInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldVarSetPageModify(ByRef $oPageVarSetField, $bRefOn = Null, $iOffset = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPage[2]

	If Not IsObj($oPageVarSetField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bRefOn, $iOffset) Then
		__LOWriter_ArrayFill($avPage, $oPageVarSetField.On(), $oPageVarSetField.Offset())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avPage)
	EndIf

	If ($bRefOn <> Null) Then
		If Not IsBool($bRefOn) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oPageVarSetField.On = $bRefOn
		$iError = ($oPageVarSetField.On() = $bRefOn) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oPageVarSetField.Offset = $iOffset
		$iError = ($oPageVarSetField.Offset() = $iOffset) ? $iError : BitOR($iError, 2)
	EndIf

	$oPageVarSetField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldVarSetPageModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldVarShowPageInsert
; Description ...: Insert a Show Page Variable Field.
; Syntax ........: _LOWriter_FieldVarShowPageInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $iNumFormat = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
;                  $iNumFormat            - [optional] an integer value. Default is Null. The numbering format to use for Show
;				   +						Page Variable numbering. See Constants.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iNumFormat not an integer, less than 0 or greater than 71. See Constants.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.ReferencePageGet" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Show Page Variable Field. Returning
;				   +											the Show Page Variable Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Numbering Format Constants: $LOW_NUM_STYLE_CHARS_UPPER_LETTER(0), Numbering is put in upper case letters. ("A, B, C, D)
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
; Related .......: _LOWriter_FieldVarShowPageModify, _LOWriter_DocGetViewCursor,	_LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldVarShowPageInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $iNumFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageShowField

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oPageShowField = $oDoc.createInstance("com.sun.star.text.TextField.ReferencePageGet")
	If Not IsObj($oPageShowField) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oPageShowField.NumberingType = $iNumFormat
	Else
		$oPageShowField.NumberingType = $LOW_NUM_STYLE_PAGE_DESCRIPTOR
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oPageShowField, $bOverwrite)

	$oPageShowField.Update()

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oPageShowField)
EndFunc   ;==>_LOWriter_FieldVarShowPageInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldVarShowPageModify
; Description ...: Set or Retrieve a Show Page Variable Field's settings.
; Syntax ........: _LOWriter_FieldVarShowPageModify(Byref $oPageShowField[, $iNumFormat = Null])
; Parameters ....: $oPageShowField      - [in/out] an object. A Show Page Variable field Object from a previous Insert or
;				   +									retrieval function.
;                  $iNumFormat            - [optional] an integer value. Default is Null. The numbering format to use for Show
;				   +						Page Variable numbering. See Constants.
; Return values .: Success: 1 or Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oPageShowField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iNumFormat not an integer, less than 0 or greater than 71. See Constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1
;				   |								1 = Error setting $iNumFormat
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current
;				   +								Numbering Type setting, in Integer format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;Numbering Format Constants: $LOW_NUM_STYLE_CHARS_UPPER_LETTER(0), Numbering is put in upper case letters. ("A, B, C, D)
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
; Related .......: _LOWriter_FieldVarShowPageInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldVarShowPageModify(ByRef $oPageShowField, $iNumFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oPageShowField) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iNumFormat) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oPageShowField.NumberingType())

	If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oPageShowField.NumberingType = $iNumFormat
	$iError = ($oPageShowField.NumberingType() = $iNumFormat) ? $iError : BitOR($iError, 1)

	$oPageShowField.Update()

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldVarShowPageModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyAlignment
; Description ...: Modify or Add Find Format Alignment Settings.
; Syntax ........: _LOWriter_FindFormatModifyAlignment(Byref $atFormat[, $iHorAlign = Null[, $iVertAlign = Null[, $iLastLineAlign = Null[, $bExpandSingleWord = Null[, $bSnapToGrid = Null[, $iTxtDirection = Null]]]]]])
; Parameters ....: $atFormat            - [in/out] an array of dll structs. A Find Format Array of Settings to modify. Array
;				   +						will be directly modified.
;                  $iHorAlign           - [optional] an integer value. Default is Null. The Horizontal alignment of the
;				   +						paragraph. See Constants below.
;                  $iVertAlign          - [optional] an integer value. Default is Null. The Vertical alignment of the paragraph.
;				   +						See Constants below. In my personal testing, searching for the Vertical Alignment
;				   +						setting using this parameter causes any results matching the searched for string to
;				   +						be replaced, whether they contain the Vert. Align format or not, this is supposed
;				   +						to be fixed in L.O. 7.6.
;                  $iLastLineAlign      - [optional] an integer value. Default is Null. Specify the alignment for the last line
;				   +						in the paragraph. See Constants below.
;                  $bExpandSingleWord   - [optional] a boolean value. Default is Null. If the last line of a justified paragraph
;				   +						consists of one word, the word is stretched to the width of the paragraph.
;                  $bSnapToGrid         - [optional] a boolean value. Default is Null. If True, Aligns the paragraph to a text
;				   +						grid (if one is active).
;                  $iTxtDirection       - [optional] an integer value. Default is Null. The Text Writing Direction. See
;				   +						Constants below. [Libre Office Default is 4] In my personal testing, searching for
;				   +						the Text Direction setting using this parameter alone, without using other
;				   +						parameters, causes any results matching the searched for string to be replaced,
;				   +						whether they contain the Text Direction format or not, this is supposed to be
;				   +						fixed in L.O. 7.6.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iHorAlign not an integer, less than 0 or greater than 3. See Constants.
;				   @Error 1 @Extended 3 Return 0 = $iVertAlign not an integer, less than 0 or more than 4. See Constants.
;				   @Error 1 @Extended 4 Return 0 = $iLastLineAlign not an integer, less than 0 or more than 3. See Constants.
;				   @Error 1 @Extended 5 Return 0 = $bExpandSingleWord not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $bSnapToGrid not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iTxtDirection not an Integer, less than 0 or greater than 5, See Constants.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
;					Note: $iTxtDirection constants 2,3, and 5 may not be available depending on your language settings.
; Horizontal Alignment Constants: $LOW_PAR_ALIGN_HOR_LEFT(0); The Paragraph is left-aligned between the borders.
;									$LOW_PAR_ALIGN_HOR_RIGHT(1); The Paragraph is right-aligned between the borders.
;									$LOW_PAR_ALIGN_HOR_JUSTIFIED(2); The Paragraph is adjusted to both borders / stretched.
;									$LOW_PAR_ALIGN_HOR_CENTER(3); The Paragraph is centered between the left and right borders.
; Vertical Alignment Constants: $LOW_PAR_ALIGN_VERT_AUTO(0); In automatic mode, horizontal text is aligned to the baseline. The
;										same applies to text that is rotated 90°. Text that is rotated 270 ° is aligned to the
;										center.
;									$LOW_PAR_ALIGN_VERT_BASELINE(1); The text is aligned to the baseline.
;									$LOW_PAR_ALIGN_VERT_TOP(2); The text is aligned to the top.
;									$LOW_PAR_ALIGN_VERT_CENTER(3); The text is aligned to the center.
;									$LOW_PAR_ALIGN_VERT_BOTTOM(4); The text is aligned to bottom.
; Last Line Alignment Constants: $LOW_PAR_LAST_LINE_START(0); The Paragraph is aligned either to the Left border or the right,
;										depending on the current text direction.
;									$LOW_PAR_LAST_LINE_JUSTIFIED(2); The Paragraph is adjusted to both borders / stretched.
;									$LOW_PAR_LAST_LINE_CENTER(3); The Paragraph is centered between the left and right borders.
; Text Direction Constants: $LOW_TXT_DIR_LR_TB(0), — text within lines is written left-to-right. Lines and blocks are placed
;								top-to-bottom. Typically, this is the writing mode for normal "alphabetic" text.
;							$LOW_TXT_DIR_RL_TB(1), — text within a line are written right-to-left. Lines and blocks are placed
;								top-to-bottom. Typically, this writing mode is used in Arabic and Hebrew text.
;							$LOW_TXT_DIR_TB_RL(2), — text within a line is written top-to-bottom. Lines and blocks are placed
;								right-to-left. Typically, this writing mode is used in Chinese and Japanese text.
;							$LOW_TXT_DIR_TB_LR(3), — text within a line is written top-to-bottom. Lines and blocks are placed
;								left-to-right. Typically, this writing mode is used in Mongolian text.
;							$LOW_TXT_DIR_CONTEXT(4), — obtain actual writing mode from the context of the object.
;							$LOW_TXT_DIR_BT_LR(5), — text within a line is written bottom-to-top. Lines and blocks are placed
;								left-to-right. (LibreOffice 6.3)
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyAlignment(ByRef $atFormat, $iHorAlign = Null, $iVertAlign = Null, $iLastLineAlign = Null, $bExpandSingleWord = Null, $bSnapToGrid = Null, $iTxtDirection = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iHorAlign <> Null) Then
		If ($iHorAlign = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaAdjust")
		Else
			If Not __LOWriter_IntIsBetween($iHorAlign, $LOW_PAR_ALIGN_HOR_LEFT, $LOW_PAR_ALIGN_HOR_CENTER) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaAdjust", $iHorAlign))
		EndIf
	EndIf

	If ($iVertAlign <> Null) Then
		If ($iVertAlign = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaVertAlignment")
		Else
			If Not __LOWriter_IntIsBetween($iVertAlign, $LOW_PAR_ALIGN_VERT_AUTO, $LOW_PAR_ALIGN_VERT_BOTTOM) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaVertAlignment", $iVertAlign))
		EndIf
	EndIf

	If ($iLastLineAlign <> Null) Then
		If ($iLastLineAlign = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaLastLineAdjust")
		Else
			If Not __LOWriter_IntIsBetween($iLastLineAlign, $LOW_PAR_LAST_LINE_JUSTIFIED, $LOW_PAR_LAST_LINE_CENTER, "", $LOW_PAR_LAST_LINE_START) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaLastLineAdjust", $iLastLineAlign))
		EndIf
	EndIf

	If ($bExpandSingleWord <> Null) Then
		If ($bExpandSingleWord = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaExpandSingleWord")
		Else
			If Not IsBool($bExpandSingleWord) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaExpandSingleWord", $bExpandSingleWord))
		EndIf
	EndIf

	If ($bSnapToGrid <> Null) Then
		If ($bSnapToGrid = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "SnapToGrid")
		Else
			If Not IsBool($bSnapToGrid) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("SnapToGrid", $bSnapToGrid))
		EndIf
	EndIf

	If ($iTxtDirection <> Null) Then
		If ($iTxtDirection = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "WritingMode")
		Else
			If Not __LOWriter_IntIsBetween($iTxtDirection, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("WritingMode", $iTxtDirection))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyEffects
; Description ...: Modify or Add Find Format Effects Settings.
; Syntax ........: _LOWriter_FindFormatModifyEffects(Byref $atFormat[,$iRelief  = Null[, $iCase = Null[, $bOutline = Null[, $bShadow = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $iRelief             - [optional] an integer value. Default is Null. The Character Relief style. See
;				   +						Constants below. Min. 0, Max 2. In my personal testing, searching for the Relief
;				   +						setting using this parameter causes any results matching the searched for string to
;				   +						be replaced, whether they contain the Relief format or not, this is supposed to be
;				   +						fixed in L.O. 7.6.
;                  $iCase               - [optional] an integer value. Default is Null. The Character Case Style. See Constants
;				   +						below. Min. 0, Max 4.
;                  $bOutline            - [optional] a boolean value. Default is Null. Whether the characters have an outline
;				   +						around the outside.
;                  $bShadow             - [optional] a boolean value. Default is Null. Whether the characters have a shadow.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iRelief not an integer or less than 0 or greater than 2. See Constants.
;				   @Error 1 @Extended 3 Return 0 = $iCase not an integer or less than 0 or greater than 4. See Constants.
;				   @Error 1 @Extended 4 Return 0 = $bOutline not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bShadow not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
; Case Constants : 	$LOW_CASEMAP_NONE(0); The case of the characters is unchanged.
;						$LOW_CASEMAP_UPPER(1); All characters are put in upper case.
;						$LOW_CASEMAP_LOWER(2); All characters are put in lower case.
;						$LOW_CASEMAP_TITLE(3); The first character of each word is put in upper case.
;						$LOW_CASEMAP_SM_CAPS(4); All characters are put in upper case, but with a smaller font height.
; Relief Constants: $LOW_RELIEF_NONE(0); no relief is used.
;						$LOW_RELIEF_EMBOSSED(1); the font relief is embossed.
;						$LOW_RELIEF_ENGRAVED(2); the font relief is engraved.
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyEffects(ByRef $atFormat, $iRelief = Null, $iCase = Null, $bOutline = Null, $bShadow = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iRelief <> Null) Then
		If ($iRelief = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharRelief")
		Else
			If Not __LOWriter_IntIsBetween($iRelief, $LOW_RELIEF_NONE, $LOW_RELIEF_ENGRAVED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharRelief", $iRelief))
		EndIf
	EndIf

	If ($iCase <> Null) Then
		If ($iCase = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharCaseMap")
		Else
			If Not __LOWriter_IntIsBetween($iCase, $LOW_CASEMAP_NONE, $LOW_CASEMAP_SM_CAPS) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharCaseMap", $iCase))
		EndIf
	EndIf

	If ($bOutline <> Null) Then
		If ($bOutline = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharContoured")
		Else
			If Not IsBool($bOutline) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharContoured", $bOutline))
		EndIf
	EndIf

	If ($bShadow <> Null) Then
		If ($bShadow = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharShadowed")
		Else
			If Not IsBool($bShadow) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharShadowed", $bShadow))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyEffects

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyFont
; Description ...: Modify or Add Find Format Font Settings.
; Syntax ........: _LOWriter_FindFormatModifyFont(Byref $oDoc, Byref $atFormat[, $sFontName = Null[, $iFontSize = Null[, $iFontWeight = Null[, $iFontPosture = Null[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified. See Remarks.
;                  $sFontName           - [optional] a string value. Default is Null. The Font name to search for.
;                  $iFontSize           - [optional] an integer value. Default is Null. The Font size to search for.
;                  $iFontWeight         - [optional] an integer value. Default is Null. The Font weight to search for. See
;				   +						Constants.
;                  $iFontPosture        - [optional] an integer value. Default is Null. The Font Posture(Italic etc.,) See
;				   +						Constants.
;                  $iFontColor          - [optional] an integer value. Default is Null. The Font Color in Long Integer format,
;				   +						Min. $LOW_COLOR_OFF(-1), Max $LOW_COLOR_WHITE(16777215). See some preset values
;				   +						in Constants below.
;                  $iTransparency       - [optional] an integer value. Default is Null. The percentage of Transparency, Min. 0,
;				   +						Max 100. 0 is not visible, 100 is fully visible. Seems to require a color entered
;				   +						in $iFontColor before transparency can be searched for. Libre Office 7.0 and Up.
;                  $iHighlight          - [optional] an integer value. Default is Null. The Highlight color to search for, in
;				   +						 Long Integer format, min. $LOW_COLOR_OFF(-1), Max $LOW_COLOR_WHITE(16777215),
;				   +						See some preset values in Color Constants below.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 3 Return 0 = $sFontName not a String.
;				   @Error 1 @Extended 4 Return 0 = Font defined in $sFontName not found in current Document.
;				   @Error 1 @Extended 5 Return 0 = $iFontSize not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iFontWeight not an Integer, less than 50 but not 0, or more than 200. See
;				   +									Constants.
;				   @Error 1 @Extended 7 Return 0 = $iFontPosture not an Integer, less than 0 or greater than 5. See Constants.
;				   @Error 1 @Extended 8 Return 0 = $iFontColor not an Integer, less than -1 or greater than 16777215.
;				   @Error 1 @Extended 9 Return 0 = $iTransparency not an Integer, Less than 0 or greater than 100.
;				   @Error 1 @Extended 10 Return 0 = $iHighlight not an Integer, less than -1 or greater than 16777215.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 7.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
; Weight Constants : $LOW_WEIGHT_DONT_KNOW(0); The font weight is not specified/known.
;						$LOW_WEIGHT_THIN(50); specifies a 50% font weight.
;						$LOW_WEIGHT_ULTRA_LIGHT(60); specifies a 60% font weight.
;						$LOW_WEIGHT_LIGHT(75); specifies a 75% font weight.
;						$LOW_WEIGHT_SEMI_LIGHT(90); specifies a 90% font weight.
;						$LOW_WEIGHT_NORMAL(100); specifies a normal font weight.
;						$LOW_WEIGHT_SEMI_BOLD(110); specifies a 110% font weight.
;						$LOW_WEIGHT_BOLD(150); specifies a 150% font weight.
;						$LOW_WEIGHT_ULTRA_BOLD(175); specifies a 175% font weight.
;						$LOW_WEIGHT_BLACK(200); specifies a 200% font weight.
; Slant/Posture Constants : $LOW_POSTURE_NONE(0); specifies a font without slant.
;							$LOW_POSTURE_OBLIQUE(1); specifies an oblique font (slant not designed into the font).
;							$LOW_POSTURE_ITALIC(2); specifies an italic font (slant designed into the font).
;							$LOW_POSTURE_DontKnow(3); specifies a font with an unknown slant.
;							$LOW_POSTURE_REV_OBLIQUE(4); specifies a reverse oblique font (slant not designed into the font).
;							$LOW_POSTURE_REV_ITALIC(5); specifies a reverse italic font (slant designed into the font).
; Color Constants: $LOW_COLOR_OFF(-1),
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
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong,_LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll _LOWriter_DocReplaceAllInRange,
;					_LOWriter_FontsList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyFont(ByRef $oDoc, ByRef $atFormat, $sFontName = Null, $iFontSize = Null, $iFontWeight = Null, $iFontPosture = Null, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $UBOUND_COLUMNS = 2

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($sFontName <> Null) Then
		If ($sFontName = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharFontName")
		Else
			If Not IsString($sFontName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			If Not _LOWriter_FontExists($oDoc, $sFontName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharFontName", $sFontName))
		EndIf
	EndIf

	If ($iFontSize <> Null) Then
		If ($iFontSize = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharHeight")
		Else
			If Not IsInt($iFontSize) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharHeight", $iFontSize))
		EndIf
	EndIf

	If ($iFontWeight <> Null) Then
		If ($iFontWeight = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharWeight")
		Else
			If Not __LOWriter_IntIsBetween($iFontWeight, $LOW_WEIGHT_THIN, $LOW_WEIGHT_BLACK, "", $LOW_WEIGHT_DONT_KNOW) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharWeight", $iFontWeight))
		EndIf
	EndIf

	If ($iFontPosture <> Null) Then
		If ($iFontPosture = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharPosture")
		Else
			If Not __LOWriter_IntIsBetween($iFontPosture, $LOW_POSTURE_NONE, $LOW_POSTURE_ITALIC) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharPosture", $iFontPosture))
		EndIf
	EndIf

	If ($iFontColor <> Null) Then
		If ($iFontColor = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharColor")
		Else
			If Not __LOWriter_IntIsBetween($iFontColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharColor", $iFontColor))
		EndIf
	EndIf

	If ($iTransparency <> Null) Then
		If ($iTransparency = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharTransparence")
		Else
			If Not __LOWriter_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
			If Not __LOWriter_VersionCheck(7.0) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharTransparence", $iTransparency))
		EndIf
	EndIf

	If ($iHighlight <> Null) Then
		If ($iHighlight = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharBackColor")
			If __LOWriter_VersionCheck(4.2) Then __LOWriter_FindFormatDeleteSetting($atFormat, "CharHighlight")
		Else
			If Not __LOWriter_IntIsBetween($iHighlight, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
			;CharHighlight; same as CharBackColor---Libre seems to use back color for highlighting.
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharBackColor", $iHighlight))
			If __LOWriter_VersionCheck(4.2) Then __LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharHighlight", $iHighlight))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyHyphenation
; Description ...: Modify or Add Find Format Hyphenation Settings. See Remarks.
; Syntax ........: _LOWriter_FindFormatModifyHyphenation(Byref $atFormat[, $bAutoHyphen = Null[, $bHyphenNoCaps = Null[, $iMaxHyphens = Null[, $iMinLeadingChar = Null[, $iMinTrailingChar = Null]]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $bAutoHyphen         - [optional] a boolean value. Default is Null. Whether  automatic hyphenation is applied.
;                  $bHyphenNoCaps       - [optional] a boolean value. Default is Null. Setting to true will disable
;				   +						hyphenation of words written in CAPS for this paragraph. Libre 6.4 and up.
;                  $iMaxHyphens         - [optional] an integer value. Default is Null. The maximum number of consecutive
;				   +						hyphens. Min 0, Max 99.
;                  $iMinLeadingChar     - [optional] an integer value. Default is Null. Specifies the minimum number of
;				   +						characters to remain before the hyphen character (when hyphenation is applied).
;				   +						Min 2, max 9.
;                  $iMinTrailingChar    - [optional] an integer value. Default is Null. Specifies the minimum number of
;				   +						characters to remain after the hyphen character (when hyphenation is applied).
;				   +						Min 2, max 9.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $bAutoHyphen not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bHyphenNoCaps not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iMaxHyphens not an Integer, less than 0, or greater than 99.
;				   @Error 1 @Extended 5 Return 0 = $iMinLeadingChar not an Integer, less than 2 or greater than 9.
;				   @Error 1 @Extended 6 Return 0 = $iMinTrailingChar not an Integer, less than 2 or greater than 9.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 6.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: In my personal testing, searching for any of these hyphenation formatting settings causes any results
;						matching the searched for string to be replaced, whether they contain these formatting settings or not,
;						I am unsure why.
;					Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyHyphenation(ByRef $atFormat, $bAutoHyphen = Null, $bHyphenNoCaps = Null, $iMaxHyphens = Null, $iMinLeadingChar = Null, $iMinTrailingChar = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($bAutoHyphen <> Null) Then
		If ($bAutoHyphen = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaIsHyphenation")
		Else
			If Not IsBool($bAutoHyphen) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaIsHyphenation", $bAutoHyphen))
		EndIf
	EndIf

	If ($bHyphenNoCaps <> Null) Then
		If ($bHyphenNoCaps = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaHyphenationNoCaps")
		Else
			If Not IsBool($bHyphenNoCaps) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			If Not __LOWriter_VersionCheck(6.4) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaHyphenationNoCaps", $bHyphenNoCaps))
		EndIf
	EndIf

	If ($iMaxHyphens <> Null) Then
		If ($iMaxHyphens = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaHyphenationMaxHyphens")
		Else
			If Not __LOWriter_IntIsBetween($iMaxHyphens, 0, 99) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaHyphenationMaxHyphens", $iMaxHyphens))
		EndIf
	EndIf

	If ($iMinLeadingChar <> Null) Then
		If ($iMinLeadingChar = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaHyphenationMaxLeadingChars")
		Else
			If Not __LOWriter_IntIsBetween($iMinLeadingChar, 2, 9) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaHyphenationMaxLeadingChars", $iMinLeadingChar))
		EndIf
	EndIf

	If ($iMinTrailingChar <> Null) Then
		If ($iMinTrailingChar = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaHyphenationMaxTrailingChars")
		Else
			If Not __LOWriter_IntIsBetween($iMinTrailingChar, 2, 9) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaHyphenationMaxTrailingChars", $iMinTrailingChar))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyHyphenation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyIndent
; Description ...: Modify or Add Find Format Indent Settings.
; Syntax ........: _LOWriter_FindFormatModifyIndent(Byref $atFormat[, $iBeforeText = Null[, $iAfterText = Null[, $iFirstLine = Null[, $bAutoFirstLine = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $iBeforeText         - [optional] an integer value. Default is Null. The amount of space that you want to
;				   +						indent the paragraph from the page margin. Set in MicroMeters(uM) Min. -9998989,
;				   +						Max.17094. Both $iBeforeText and $iAfterText must be set to perform a search for
;				   +						either.
;                  $iAfterText          - [optional] an integer value. Default is Null. The amount of space that you want to
;				   +						indent the paragraph from the page margin. Set in MicroMeters(uM) Min. -9998989,
;				   +						Max.17094. Both $iBeforeText and $iAfterText must be set to perform a search for
;				   +						either.
;                  $iFirstLine          - [optional] an integer value. Default is Null. Indentation distance of the first line
;				   +						of a paragraph, Set in MicroMeters(uM) Min. -57785, Max.17094. Both $iBeforeText and
;				   +						$iAfterText must be set to perform a search for $iFirstLine.
;                  $bAutoFirstLine      - [optional] a boolean value. Default is Null. Whether the first line should be indented
;				   +						automatically. Both $iBeforeText and $iAfterText must be set to perform a search
;				   +						for $bAutoFirstLine.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iBeforeText not an integer, less than -9998989 or more than 17094 uM.
;				   @Error 1 @Extended 3 Return 0 = $iAfterText not an integer, less than -9998989 or more than 17094 uM.
;				   @Error 1 @Extended 4 Return 0 = $iFirstLine not an integer, less than -57785 or more than 17094 uM.
;				   @Error 1 @Extended 5 Return 0 = $bAutoFirstLine not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
;					Note: $iFirstLine Indent cannot be set if $bAutoFirstLine is set to True.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyIndent(ByRef $atFormat, $iBeforeText = Null, $iAfterText = Null, $iFirstLine = Null, $bAutoFirstLine = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	;Min: -9998989;Max: 17094
	If ($iBeforeText <> Null) Then
		If ($iBeforeText = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaLeftMargin")
		Else
			If Not __LOWriter_IntIsBetween($iBeforeText, -9998989, 17094) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaLeftMargin", $iBeforeText))
		EndIf
	EndIf

	If ($iAfterText <> Null) Then
		If ($iAfterText = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaRightMargin")
		Else
			If Not __LOWriter_IntIsBetween($iAfterText, -9998989, 17094) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaRightMargin", $iAfterText))
		EndIf
	EndIf

	;mx 17094min;-57785
	If ($iFirstLine <> Null) Then
		If ($iFirstLine = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaFirstLineIndent")
		Else
			If Not __LOWriter_IntIsBetween($iFirstLine, -57785, 17094) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaFirstLineIndent", $iFirstLine))
		EndIf
	EndIf

	If ($bAutoFirstLine <> Null) Then
		If ($bAutoFirstLine = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaIsAutoFirstLineIndent")
		Else
			If Not IsBool($bAutoFirstLine) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaIsAutoFirstLineIndent", $bAutoFirstLine))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyIndent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyOverline
; Description ...: Modify or Add Find Format Overline Settings.
; Syntax ........: _LOWriter_FindFormatModifyOverline(Byref $atFormat[, $iOverLineStyle = Null[, $bWordOnly = Null[, $bOLHasColor = Null[, $iOLColor = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $iOverLineStyle      - [optional] an integer value. Default is Null. The style of the Overline line, see
;				   +						constants listed below, Underline Constants are used for Overline line style.
;				   +						Overline style must be set before any of the other parameters can be searched for.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not Overlined.
;				   +						See remarks.
;                  $bOLHasColor         - [optional] a boolean value. Default is Null. Whether the Overline is colored, must
;				   +						be set to true in order to set the Overline color.
;                  $iOLColor            - [optional] an integer value. Default is Null. The color of the Overline, set in Long
;				   +						integer format. Can be one of the constants below or a custom value.
;				   +						$LOW_COLOR_OFF(-1) is automatic color mode.  Min. -1, Max 16777215.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iOverLineStyle not an Integer, or less than 0 or greater than 18. See
;				   +									Constants.
;				   @Error 1 @Extended 3 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bOLHasColor not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iOLColor not an Integer, or less than -1 or greater than 16777215.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
;					$bWordOnly applies to Underline, Overline and Strikeout, regardless of which is set to true, one setting
;						applies to all.
; OverLine line style Constants: $LOW_UNDERLINE_NONE(0),
;									$LOW_UNDERLINE_SINGLE(1),
;									$LOW_UNDERLINE_DOUBLE(2),
;									$LOW_UNDERLINE_DOTTED(3),
;									$LOW_UNDERLINE_DONT_KNOW(4),
;									$LOW_UNDERLINE_DASH(5),
;									$LOW_UNDERLINE_LONG_DASH(6),
;									$LOW_UNDERLINE_DASH_DOT(7),
;									$LOW_UNDERLINE_DASH_DOT_DOT(8),
;									$LOW_UNDERLINE_SML_WAVE(9),
;									$LOW_UNDERLINE_WAVE(10),
;									$LOW_UNDERLINE_DBL_WAVE(11),
;									$LOW_UNDERLINE_BOLD(12),
;									$LOW_UNDERLINE_BOLD_DOTTED(13),
;									$LOW_UNDERLINE_BOLD_DASH(14),
;									$LOW_UNDERLINE_BOLD_LONG_DASH(15),
;									$LOW_UNDERLINE_BOLD_DASH_DOT(16),
;									$LOW_UNDERLINE_BOLD_DASH_DOT_DOT(17),
;									$LOW_UNDERLINE_BOLD_WAVE(18)
; Color Constants: $LOW_COLOR_OFF(-1),
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
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong, _LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyOverline(ByRef $atFormat, $iOverLineStyle = Null, $bWordOnly = Null, $bOLHasColor = Null, $iOLColor = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iOverLineStyle <> Null) Then
		If ($iOverLineStyle = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharOverline")
		Else
			If Not __LOWriter_IntIsBetween($iOverLineStyle, $LOW_UNDERLINE_NONE, $LOW_UNDERLINE_BOLD_WAVE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharOverline", $iOverLineStyle))
		EndIf
	EndIf

	If ($bWordOnly <> Null) Then
		If ($bWordOnly = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharWordMode")
		Else
			If Not IsBool($bWordOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharWordMode", $bWordOnly))
		EndIf
	EndIf

	If ($bOLHasColor <> Null) Then
		If ($bOLHasColor = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharOverlineHasColor")
		Else
			If Not IsBool($bOLHasColor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharOverlineHasColor", $bOLHasColor))
		EndIf
	EndIf

	If ($iOLColor <> Null) Then
		If ($iOLColor = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharOverlineColor")
		Else
			If Not __LOWriter_IntIsBetween($iOLColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharOverlineColor", $iOLColor))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyOverline

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyPageBreak
; Description ...: Modify or Add Find Format Page Break Settings. See Remarks.
; Syntax ........: _LOWriter_FindFormatModifyPageBreak(Byref $oDoc, Byref $atFormat[, $iBreakType = Null[, $sPageStyle = Null[, $iPgNumOffSet = Null]]])
; Parameters ....: $oDoc                - [in/out] an object.  A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $iBreakType          - [optional] an integer value. Default is Null. The Page Break Type. See Constants below.
;                  $sPageStyle          - [optional] a string value. Default is Null. Creates a page break before the paragraph
;				   +						it belongs to and assigns the value as the name of the new page style to use.
;                  $iPgNumOffSet        - [optional] an integer value. Default is Null. If a page break property is set at a
;				   +						paragraph, this property contains the new value for the page number.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iBreakType not an integer, less than 0 or greater than 6.
;				   @Error 1 @Extended 3 Return 0 = $sPageStyle not a String.
;				   @Error 1 @Extended 4 Return 0 = Page Style defined in $sPageStyle not found in current document.
;				   @Error 1 @Extended 5 Return 0 = $iPgNumOffSet not an Integer or less than 0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: In my personal testing, searching for a page break was very hit and miss, especially when searching with the
;					"PageStyle" Name parameter, and it never worked for searching for PageNumberOffset.
;					Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
;Break Constants : $LOW_BREAK_NONE(0) – No column or page break is applied.
;						$LOW_BREAK_COLUMN_BEFORE(1) – A column break is applied before the current Paragraph. The current
;							Paragraph, therefore, is the first in the column.
;						$LOW_BREAK_COLUMN_AFTER(2) – A column break is applied after the current Paragraph. The current
;							Paragraph, therefore, is the last in the column.
;						$LOW_BREAK_COLUMN_BOTH(3) – A column break is applied before and after the current Paragraph. The
;							current Paragraph, therefore, is the only Paragraph in the column.
;						$LOW_BREAK_PAGE_BEFORE(4) – A page break is applied before the current Paragraph. The current Paragraph,
;						therefore, is the first on the page.
;						$LOW_BREAK_PAGE_AFTER(5) – A page break is applied after the current Paragraph. The current Paragraph,
;						therefore, is the last on the page.
;						$LOW_BREAK_PAGE_BOTH(6) – A page break is applied before and after the current Paragraph. The current
;						Paragraph, therefore, is the only paragraph on the page.
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyPageBreak(ByRef $oDoc, ByRef $atFormat, $iBreakType = Null, $sPageStyle = Null, $iPgNumOffSet = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iBreakType <> Null) Then
		If ($iBreakType = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "BreakType")
		Else
			If Not __LOWriter_IntIsBetween($iBreakType, $LOW_BREAK_NONE, $LOW_BREAK_PAGE_BOTH) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("BreakType", $iBreakType))
		EndIf
	EndIf

	If ($sPageStyle <> Null) Then
		If ($sPageStyle = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "PageStyleName") ;PageDescName -- Not working?
		Else
			If Not IsString($sPageStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			If Not _LOWriter_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("PageStyleName", $sPageStyle))
		EndIf
	EndIf

	If ($iPgNumOffSet <> Null) Then
		If ($iPgNumOffSet = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "PageNumberOffset")
		Else
			If Not __LOWriter_IntIsBetween($iPgNumOffSet, 0, $iPgNumOffSet) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("PageNumberOffset", $iPgNumOffSet))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyPageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyPosition
; Description ...: Modify or Add Find Format Position Settings.
; Syntax ........: _LOWriter_FindFormatModifyPosition(Byref $atFormat[, $bAutoSuper = Null[, $iSuperScript = Null[, $bAutoSub = Null[, $iSubScript = Null[, $iRelativeSize = Null]]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $bAutoSuper          - [optional] a boolean value. Default is Null.  Whether to active automatic sizing for
;				   +						SuperScript. Note: $iRelativeSize must be set to be able to search for
;				   +						Super/SubScript settings.
;                  $iSuperScript        - [optional] an integer value. Default is Null. SuperScript percentage value. See
;				   +						Remarks. Note: $iRelativeSize must be set to be able to search for Super/SubScript
;				   +						settings.
;                  $bAutoSub            - [optional] a boolean value. Default is Null. Whether to active automatic sizing for
;				   +						SubScript. Note: $iRelativeSize must be set to be able to search for Super/SubScript
;				   +						settings.
;                  $iSubScript          - [optional] an integer value. Default is Null. SubScript percentage value. See Remarks.
;				   +						Note: $iRelativeSize must be set to be able to search for Super/SubScript settings.
;                  $iRelativeSize       - [optional] an integer value. Default is Null. Percentage relative to current font size,
;				   +						Min. 1, Max 100.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $bAutoSuper not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bAutoSub not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iSuperScript not an integer, or less than 0, higher than 100 and Not 14000.
;				   @Error 1 @Extended 5 Return 0 = $iSubScript not an integer, or less than -100, higher than 100 and Not
;				   +									(-)14000.
;				   @Error 1 @Extended 6 Return 0 = $iRelativeSize not an integer, or less than 1, higher than 100.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
;					0 is the normal $iSubScript or $iSuperScript setting.
;					The way LibreOffice is set up Super/SubScript are set in the same setting, Super is a positive number from
;						1 to 100 (percentage), SubScript is a negative number set to 1 to 100 percentage. For the user's
;						convenience this function accepts both positive and negative numbers for SubScript, if a positive number
;						is called for SubScript, it is automatically set to a negative. Automatic Superscript has a integer
;						value of 14000, Auto SubScript has a integer value of -14000. There is no settable setting of Automatic
;						Super/Sub Script, though one exists, it is read-only in LibreOffice, consequently I have made two
;						separate parameters to be able to determine if the user wants to automatically set SuperScript or
;						SubScript. If you set both Auto SuperScript to True and Auto SubScript to True, or $iSuperScript to an
;						integer and $iSubScript to an integer, Subscript will be set as it is the last in the line to be set in
;						this function, and thus will over-write any SuperScript settings.
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyPosition(ByRef $atFormat, $bAutoSuper = Null, $iSuperScript = Null, $bAutoSub = Null, $iSubScript = Null, $iRelativeSize = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($bAutoSuper <> Null) Then
		If ($bAutoSuper = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharEscapement")
		Else
			If Not IsBool($bAutoSuper) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			;If $bAutoSuper = True set it to 14000 (automatic superScript) else if $iSuperScript is set, let that overwrite
			;	the current setting, else if subscript is true or set to an integer, it will overwrite the setting. If nothing
			;else set SubScript to 1
			$iSuperScript = ($bAutoSuper) ? 14000 : (IsInt($iSuperScript)) ? $iSuperScript : (IsInt($iSubScript) Or ($bAutoSub = True)) ? $iSuperScript : 1
		EndIf
	EndIf

	If ($bAutoSub <> Null) Then
		If ($bAutoSub = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharEscapement")
		Else
			If Not IsBool($bAutoSub) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			;If $bAutoSub = True set it to -14000 (automatic SubScript) else if $iSubScript is set, let that overwrite
			;	the current setting, else if superscript is true or set to an integer, it will overwrite the setting.
			$iSubScript = ($bAutoSub) ? -14000 : (IsInt($iSubScript)) ? $iSubScript : (IsInt($iSuperScript)) ? $iSubScript : 1
		EndIf
	EndIf

	If ($iSuperScript <> Null) Then
		If ($iSuperScript = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharEscapement")
		Else
			If Not __LOWriter_IntIsBetween($iSuperScript, 0, 100, "", 14000) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharEscapement", $iSuperScript))
		EndIf
	EndIf

	If ($iSubScript <> Null) Then
		If ($iSubScript = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharEscapement")
		Else
			If Not __LOWriter_IntIsBetween($iSubScript, -100, 100, "", "-14000:14000") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			$iSubScript = ($iSubScript > 0) ? Int("-" & $iSubScript) : $iSubScript
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharEscapement", $iSubScript))
		EndIf
	EndIf

	If ($iRelativeSize <> Null) Then
		If ($iRelativeSize = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharEscapementHeight")
		Else
			If Not __LOWriter_IntIsBetween($iRelativeSize, 1, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharEscapementHeight", $iRelativeSize))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyRotateScaleSpace
; Description ...: Modify or Add Find Format Rotate, Scale, and Space Settings.
; Syntax ........: _LOWriter_FindFormatModifyRotateScaleSpace(Byref $atFormat[, $iRotation = Null[, $iScaleWidth = Null[, $bAutoKerning = Null[, $nKerning = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $iRotation           - [optional] an integer value. Default is Null. Degrees to rotate the text. Accepts only
;				   +						0, 90, and 270 degrees. In my personal testing, searching for the Rotate
;				   +						setting using this parameter causes any results matching the searched for string to
;				   +						be replaced, whether they contain the Rotate format or not, this is supposed to be
;				   +						fixed in L.O. 7.6.
;                  $iScaleWidth         - [optional] an integer value. Default is Null. The percentage to horizontally stretch
;				   +						or compress the text. Min. 1. Max 100. 100 is normal sizing. In my personal testing,
;				   +						searching for the Scale Width setting using this parameter causes any results
;				   +						matching the searched for string to be replaced, whether they contain the Scale
;				   +						Width format or not, this is supposed to be fixed in L.O. 7.6.
;                  $bAutoKerning        - [optional] a boolean value. Default is Null. True applies a spacing in between certain
;				   +						pairs of characters. False = disabled.
;                  $nKerning            - [optional] a general number value. Default is Null. The kerning value of the
;				   +						characters. Min is -2 Pt. Max is 928.8 Pt. See Remarks. Values are in Printer's
;				   +						Points as set in the Libre Office UI.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iRotation not an Integer or not equal to 0, 90 or 270 degrees.
;				   @Error 1 @Extended 3 Return 0 = $iScaleWidth not an Integer or less than 1 or greater than 100.
;				   @Error 1 @Extended 4 Return 0 = $bAutoKerning not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $nKerning not a number, or less than -2 or greater than 928.8 Points.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
;					When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User
;						Display, however the internal setting is measured in MicroMeters. They will be automatically converted
;						from Points to MicroMeters and back for retrieval of settings.
;						The acceptable values are from -2 Pt to  928.8 Pt. the figures can be directly converted easily,
;						however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative
;						MicroMeters internally from 928.9 up to 1000 Pt (Max setting). For example, 928.8Pt is the last
;						correct value, which equals 32766 uM (MicroMeters), after this LibreOffice reports the following:
;						;928.9 Pt = -32766 uM; 929 Pt = -32763 uM; 929.1 = -32759; 1000 pt = -30258. Attempting to set
;						Libre's kerning value to anything over 32768 uM causes a COM exception, and attempting to set the
;						 kerning to any of these negative numbers sets the User viewable kerning value to -2.0 Pt. For these
;						reasons the max settable kerning is -2.0 Pt to 928.8 Pt.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyRotateScaleSpace(ByRef $atFormat, $iRotation = Null, $iScaleWidth = Null, $bAutoKerning = Null, $nKerning = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iRotation <> Null) Then
		If ($iRotation = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharRotation")
		Else
			If Not __LOWriter_IntIsBetween($iRotation, 0, 0, "", "90:270") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			$iRotation = ($iRotation > 0) ? ($iRotation * 10) : $iRotation ;rotation set in hundredths (90 deg = 900 etc)so times by 10.
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharRotation", $iRotation))
		EndIf
	EndIf

	If ($iScaleWidth <> Null) Then ;can't be less than 1%
		If ($iScaleWidth = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharScaleWidth")
		Else
			If Not __LOWriter_IntIsBetween($iScaleWidth, 1, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharScaleWidth", $iScaleWidth))
		EndIf
	EndIf

	If ($bAutoKerning <> Null) Then
		If ($bAutoKerning = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharAutoKerning")
		Else
			If Not IsBool($bAutoKerning) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharAutoKerning", $bAutoKerning))
		EndIf
	EndIf

	If ($nKerning <> Null) Then
		If ($nKerning = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharKerning")
		Else
			If Not __LOWriter_NumIsBetween($nKerning, -2, 928.8) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			$nKerning = __LOWriter_UnitConvert($nKerning, $__LOWCONST_CONVERT_PT_UM)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharKerning", $nKerning))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyRotateScaleSpace

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifySpacing
; Description ...: Modify or Add Find Format Spacing Settings.
; Syntax ........: _LOWriter_FindFormatModifySpacing(Byref $atFormat[, $iAbovePar = Null[, $iBelowPar = Null[, $bAddSpace = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null]]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $iAbovePar           - [optional] an integer value. Default is Null. The Space above a paragraph, in
;				   +						Micrometers. Min 0 Micrometers (uM) Max 10,008 uM.
;                  $iBelowPar           - [optional] an integer value. Default is Null. The Space Below a paragraph, in
;				   +						Micrometers. Min 0, Max 10,008 Micrometers (uM).
;                  $bAddSpace           - [optional] a boolean value. Default is Null. If true, the top and bottom margins of
;				   +						the paragraph should not be applied when the previous and next paragraphs have the
;				   +						same style. Libre Office version 3.6 and up.
;                  $iLineSpcMode        - [optional] an integer value. Default is Null. The type of the line spacing of a
;				   +						paragraph. See Constants below, also notice min and max values for each. Must set
;				   +						both $iLineSpcMode and $iLineSpcHeight to be able to search either.
;                  $iLineSpcHeight      - [optional] an integer value. Default is Null. This value specifies the spacing
;				   +						of the lines. See Remarks for Minimum and Max values. Must set both $iLineSpcMode
;				   +						and $iLineSpcHeight to be able to search either.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iAbovePar not an integer, less than 0 or more than 10008 uM.
;				   @Error 1 @Extended 3 Return 0 = $iBelowPar not an integer, less than 0 or more than 10008 uM.
;				   @Error 1 @Extended 4 Return 0 = $bAddSpace not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iLineSpcMode Not an integer, less than 0 or greater than 3. See Constants.
;				   @Error 1 @Extended 6 Return 0 = $iLineSpcHeight Not an integer.
;				   @Error 1 @Extended 7 Return 0 = $iLineSpcMode set to 0(Proportional) and $iLineSpcHeight less than 6(%)
;				   +									or greater than 65535(%).
;				   @Error 1 @Extended 8 Return 0 = $iLineSpcMode set to 1 or 2(Minimum, or Leading) and $iLineSpcHeight less
;				   +								than 0 uM or greater than 10008 uM
;				   @Error 1 @Extended 9 Return 0 = $iLineSpcMode set to 3(Fixed) and $iLineSpcHeight less than 51 uM
;				   +									or greater than 10008 uM.
;				   --Initialization Errors--
;				   @Error 2 @Extended 2 Return 0 = Error creating LineSpacing Object.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
;					Note: The settings in Libre Office, (Single,1.15, 1.5, Double,) Use the Proportional mode, and are just
;						varying percentages. e.g Single = 100, 1.15 = 115%, 1.5 = 150%, Double = 200%.
;					$iLineSpcHeight depends on the $iLineSpcMode used, see constants for accepted Input values.
;					Note: $iAbovePar, $iBelowPar, $iLineSpcHeight may change +/- 1 MicroMeter once set.
; Spacing Constants :$LOW_LINE_SPC_MODE_PROP(0); This specifies the height value as a proportional value. Min 6% Max 65,535%.
;							(without percentage sign)
;						$LOW_LINE_SPC_MODE_MIN(1); (Minimum/At least) This specifies the height as the minimum line height.
;							Min 0, Max 10008 MicroMeters (uM)
;						$LOW_LINE_SPC_MODE_LEADING(2); This specifies the height value as the distance to the previous line.
;							Min 0, Max 10008 MicroMeters (uM)
;						$LOW_LINE_SPC_MODE_FIX(3); This specifies the height value as a fixed line height. Min 51 MicroMeters,
;							Max 10008 MicroMeters (uM)
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll _LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifySpacing(ByRef $atFormat, $iAbovePar = Null, $iBelowPar = Null, $bAddSpace = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tLine
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iAbovePar <> Null) Then
		If ($iAbovePar = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaTopMargin")
		Else
			If Not __LOWriter_IntIsBetween($iAbovePar, 0, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaTopMargin", $iAbovePar))
		EndIf
	EndIf

	If ($iBelowPar <> Null) Then
		If ($iBelowPar = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaBottomMargin")
		Else
			If Not __LOWriter_IntIsBetween($iBelowPar, 0, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaBottomMargin", $iBelowPar))
		EndIf
	EndIf

	If ($bAddSpace <> Null) Then
		If ($bAddSpace = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaContextMargin")
		Else
			If Not IsBool($bAddSpace) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			If Not __LOWriter_VersionCheck(3.6) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaContextMargin", $bAddSpace))
		EndIf
	EndIf

	If ($iLineSpcMode <> Null) Or ($iLineSpcHeight <> Null) Then
		If ($iLineSpcMode = Default) Or ($iLineSpcHeight = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaLineSpacing")
		Else
			$tLine = __LOWriter_FindFormatRetrieveSetting($atFormat, "ParaLineSpacing") ;Retrieve the ParaLineSpacing Property to modify if it exists.
			If (@error = 0) And (@extended = 1) Then $tLine = $tLine.Value() ;If retrieval was successful, obtain the Line Space Structure.
			If Not IsObj($tLine) Then $tLine = __LOWriter_CreateStruct("com.sun.star.style.LineSpacing") ;If retrieval was not successful, then create a new one.
			If Not IsObj($tLine) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

			If ($iLineSpcMode <> Default) And ($iLineSpcMode <> Null) Then
				If Not __LOWriter_IntIsBetween($iLineSpcMode, $LOW_LINE_SPC_MODE_PROP, $LOW_LINE_SPC_MODE_FIX) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
				$tLine.Mode = $iLineSpcMode
			EndIf

			If ($iLineSpcHeight <> Default) And ($iLineSpcHeight <> Null) Then
				If Not IsInt($iLineSpcHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
				Switch $tLine.Mode()
					Case $LOW_LINE_SPC_MODE_PROP ;Proportional
						If Not __LOWriter_IntIsBetween($iLineSpcHeight, 6, 65535) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0) ; Min setting on Proportional is 6%
					Case $LOW_LINE_SPC_MODE_MIN, $LOW_LINE_SPC_MODE_LEADING ;Minimum and Leading Modes
						If Not __LOWriter_IntIsBetween($iLineSpcHeight, 0, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
					Case $LOW_LINE_SPC_MODE_FIX ;Fixed Line Spacing Mode
						If Not __LOWriter_IntIsBetween($iLineSpcHeight, 51, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0) ; Min spacing is 51 when Fixed Mode
				EndSwitch
				$tLine.Height = $iLineSpcHeight
			EndIf


			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaLineSpacing", $tLine))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifySpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyStrikeout
; Description ...: Modify or Add Find Format Strikeout Settings.
; Syntax ........: _LOWriter_FindFormatModifyStrikeout(Byref $atFormat[, $bWordOnly = Null[, $bStrikeOut = Null[, $iStrikeLineStyle = Null]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not Overlined.
;				   +						See remarks.
;                  $bStrikeOut          - [optional] a boolean value. Default is Null. True = strikeout, False = no strike out.
;                  $iStrikeLineStyle    - [optional] an integer value. Default is Null. The Strikeout Line Style, see constants
;				   +						below.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bStrikeOut not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iStrikeLineStyle not an Integer, or less than 0 or greater than 8. See
;				   +									Constants.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
;					$bWordOnly applies to Underline, Overline and Strikeout, regardless of which is set to true, one setting
;						applies to all.
; Strikeout Line Style Constants : $LOW_STRIKEOUT_NONE(0); specifies not to strike out the characters.
;					$LOW_STRIKEOUT_SINGLE(1); specifies to strike out the characters with a single line
;					$LOW_STRIKEOUT_DOUBLE(2); specifies to strike out the characters with a double line.
;					$LOW_STRIKEOUT_DONT_KNOW(3); The strikeout mode is not specified.
;					$LOW_STRIKEOUT_BOLD(4); specifies to strike out the characters with a bold line.
;					$LOW_STRIKEOUT_SLASH(5); specifies to strike out the characters with slashes.
;					$LOW_STRIKEOUT_X(6); specifies to strike out the characters with X's.
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyStrikeout(ByRef $atFormat, $bWordOnly = Null, $bStrikeOut = Null, $iStrikeLineStyle = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($bWordOnly <> Null) Then
		If ($bWordOnly = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharWordMode")
		Else
			If Not IsBool($bWordOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharWordMode", $bWordOnly))
		EndIf
	EndIf

	If ($bStrikeOut <> Null) Then
		If ($bStrikeOut = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharCrossedOut")
		Else
			If Not IsBool($bStrikeOut) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharCrossedOut", $bStrikeOut))
		EndIf
	EndIf

	If ($iStrikeLineStyle <> Null) Then
		If ($iStrikeLineStyle = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharStrikeout")
		Else
			If Not __LOWriter_IntIsBetween($iStrikeLineStyle, $LOW_STRIKEOUT_NONE, $LOW_STRIKEOUT_X) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharStrikeout", $iStrikeLineStyle))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyStrikeout

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyTxtFlowOpt
; Description ...: Modify or Add Find Format Text Flow Settings.
; Syntax ........: _LOWriter_FindFormatModifyTxtFlowOpt(Byref $atFormat[, $bParSplit = Null[, $bKeepTogether = Null[, $iParOrphans = Null[, $iParWidows = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $bParSplit           - [optional] a boolean value. Default is Null. FALSE prevents the paragraph from getting
;				   +						split into two pages or columns
;                  $bKeepTogether       - [optional] a boolean value. Default is Null. TRUE prevents page or column breaks
;				   +						between this and the following paragraph.
;                  $iParOrphans         - [optional] an integer value. Default is Null. Specifies the minimum number of lines of
;				   +						the paragraph that have to be at bottom of a page if the paragraph is spread over
;				   +						more than one page. Min is 0 (disabled), and cannot be 1. Max is 9. In my personal
;				   +						testing, searching for the Orphan setting using this parameter causes any results
;				   +						matching the searched for string to be replaced, whether they contain the Orphan
;				   +						format or not, I am unsure why.
;                  $iParWidows          - [optional] an integer value. Default is Null. Specifies the minimum number of lines of
;				   +						the paragraph that have to be at top of a page if the paragraph is spread over more
;				   +						than one page. Min is 0 (disabled), and cannot be 1. Max is 9. In my personal
;				   +						testing, searching for the Widow setting using this parameter causes any results
;				   +						matching the searched for string to be replaced, whether they contain the Widow
;				   +						format or not, I am unsure why.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $bParSplit not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bKeepTogether not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iParOrphans not an Integer, less than 0, equal to 1, or greater than 9.
;				   @Error 1 @Extended 5 Return 0 = $iParWidows not an Integer, less than 0, equal to 1, or greater than 9.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyTxtFlowOpt(ByRef $atFormat, $bParSplit = Null, $bKeepTogether = Null, $iParOrphans = Null, $iParWidows = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($bParSplit <> Null) Then
		If ($bParSplit = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaSplit")
		Else
			If Not IsBool($bParSplit) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaSplit", $bParSplit))
		EndIf
	EndIf

	If ($bKeepTogether <> Null) Then
		If ($bKeepTogether = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaKeepTogether")
		Else
			If Not IsBool($bKeepTogether) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaKeepTogether", $bKeepTogether))
		EndIf
	EndIf

	If ($iParOrphans <> Null) Then
		If ($iParOrphans = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaOrphans")
		Else
			If Not __LOWriter_IntIsBetween($iParOrphans, 0, 9, 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaOrphans", $iParOrphans))
		EndIf
	EndIf

	If ($iParWidows <> Null) Then
		If ($iParWidows = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaWidows")
		Else
			If Not __LOWriter_IntIsBetween($iParWidows, 0, 9, 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaWidows", $iParWidows))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyTxtFlowOpt

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyUnderline
; Description ...: Modify or Add Find Format Underline Settings.
; Syntax ........: _LOWriter_FindFormatModifyUnderline(Byref $atFormat[, $iUnderLineStyle = Null[, $bWordOnly = Null[, $bULHasColor = Null[, $iULColor = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will
;				   +						 be directly modified.
;                  $iUnderLineStyle     - [optional] an integer value. Default is Null. The style of the Underline line, see
;				   +						constants listed below. Underline style must be set before any of the other
;				   +						parameters can be searched for.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not underlined.
;				   +						See remarks.
;                  $bULHasColor         - [optional] a boolean value. Default is Null. Whether the underline is colored, must
;				   +						be set to true in order to set the underline color.
;                  $iULColor            - [optional] an integer value. Default is Null. The color of the underline, set in Long
;				   +						integer format. Can be one of the constants below or a custom value.
;				   +						$LOW_COLOR_OFF(-1) is automatic color mode. Min. -1, Max 16777215.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iUnderLineStyle not an Integer, or less than 0 or greater than 18. See
;				   +									Constants.
;				   @Error 1 @Extended 3 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bULHasColor not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iULColor not an Integer, or less than -1 or greater than 16777215.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local
;						$aArray[0])
;					$bWordOnly applies to Underline, Overline and Strikeout, regardless of which is set to true, one setting
;						applies to all.
; UnderLine line style Constants: $LOW_UNDERLINE_NONE(0),
;									$LOW_UNDERLINE_SINGLE(1),
;									$LOW_UNDERLINE_DOUBLE(2),
;									$LOW_UNDERLINE_DOTTED(3),
;									$LOW_UNDERLINE_DONT_KNOW(4),
;									$LOW_UNDERLINE_DASH(5),
;									$LOW_UNDERLINE_LONG_DASH(6),
;									$LOW_UNDERLINE_DASH_DOT(7),
;									$LOW_UNDERLINE_DASH_DOT_DOT(8),
;									$LOW_UNDERLINE_SML_WAVE(9),
;									$LOW_UNDERLINE_WAVE(10),
;									$LOW_UNDERLINE_DBL_WAVE(11),
;									$LOW_UNDERLINE_BOLD(12),
;									$LOW_UNDERLINE_BOLD_DOTTED(13),
;									$LOW_UNDERLINE_BOLD_DASH(14),
;									$LOW_UNDERLINE_BOLD_LONG_DASH(15),
;									$LOW_UNDERLINE_BOLD_DASH_DOT(16),
;									$LOW_UNDERLINE_BOLD_DASH_DOT_DOT(17),
;									$LOW_UNDERLINE_BOLD_WAVE(18)
; Color Constants: $LOW_COLOR_OFF(-1),
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
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong, _LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyUnderline(ByRef $atFormat, $iUnderLineStyle = Null, $bWordOnly = Null, $bULHasColor = Null, $iULColor = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iUnderLineStyle <> Null) Then
		If ($iUnderLineStyle = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharUnderline")
		Else
			If Not __LOWriter_IntIsBetween($iUnderLineStyle, $LOW_UNDERLINE_NONE, $LOW_UNDERLINE_BOLD_WAVE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharUnderline", $iUnderLineStyle))
		EndIf
	EndIf

	If ($bWordOnly <> Null) Then
		If ($bWordOnly = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharWordMode")
		Else
			If Not IsBool($bWordOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharWordMode", $bWordOnly))
		EndIf
	EndIf

	If ($bULHasColor <> Null) Then
		If ($bULHasColor = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharUnderlineHasColor")
		Else
			If Not IsBool($bULHasColor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharUnderlineHasColor", $bULHasColor))
		EndIf
	EndIf

	If ($iULColor <> Null) Then
		If ($iULColor = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharUnderlineColor")
		Else
			If Not __LOWriter_IntIsBetween($iULColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharUnderlineColor", $iULColor))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyUnderline

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FontExists
; Description ...: Tests whether a Document has a specific font available by name.
; Syntax ........: _LOWriter_FontExists(Byref $oDoc, $sFontName)
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sFontName           - a string value. The Font name to search for.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFontName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Font list.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean  = Success. Returns True if Font is contained in the document, else
;				   +										False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function may cause a processor usage spike for a moment or two. If you wish to eliminate this, comment
;						out the current sleep function and place a sleep(10) in its place.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteDelete
; Description ...: Delete a Footnote.
; Syntax ........: _LOWriter_FootnoteDelete(Byref $oFootNote)
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous Footnote insert, or retrieval
;				   +							function.
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
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous Footnote insert, or retrieval
;				   +							function.
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
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous Footnote insert, or retrieval
;				   +							function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation
;				   +						Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the
;				   +									cursor will be overwritten. If False, content will be inserted to the
;				   +									left of any selection.
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
; Related .......: _LOWriter_FootnoteDelete, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor
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
			Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0) ;Unsupported cursor type.
		Case $LOW_CURDATA_BODY_TEXT, $LOW_CURDATA_CELL
			$oFootNote = $oDoc.createInstance("com.sun.star.text.Footnote")
			If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

		Case Else
			Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0) ;Unknown Cursor type.
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
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous Footnote insert, or retrieval
;				   +							function.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the Footnote. Set
;				   +							to "" for automatic numbering.
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
;				   @Error 0 @Extended 2 Return String = Success. $sLabel set to Null, current Footnote AutoNumbering number
;				   +									returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
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
		;If Label is blank, return the AutoNumbering Number.
		If ($oFootNote.Label() = "") Then Return SetError($__LOW_STATUS_SUCCESS, 2, $oFootNote.Anchor.String())

		;Else return the Label.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +						DocCreate function.
;                  $iNumFormat            - [optional] an integer value. Default is Null. The numbering format to use for
;				   +						Footnote numbering. See Constants.
;                  $iStartAt            - [optional] an integer value. Default is Null. The Number to begin Footnote counting
;				   +							from, this is labeled "Counting" in the L.O. User Interface. Min. 1, Max 9999.
;                  $sBefore             - [optional] a string value. Default is Null. The text to display before a Footnote
;				   +							number in the note text.
;                  $sAfter              - [optional] a string value. Default is Null. The text to display after a Footnote
;				   +							number in the note text.
;                  $iCounting           - [optional] an integer value. Default is Null. The Counting type of the footnotes,
;				   +							such as per page etc., see constants below.
;                  $bEndOfDoc           - [optional] a boolean value. Default is Null. If True, Footnotes are placed at the
;				   +							end of the document, like Endnotes.
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
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8, 16, 32
;				   |								1 = Error setting $iNumFormat
;				   |								2 = Error setting $iStartAt
;				   |								4 = Error setting $sBefore
;				   |								8 = Error setting $sAfter
;				   |								16 = Error setting $iCounting
;				   |								32 = Error setting $bEndOfDoc
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;Numbering Format Constants: $LOW_NUM_STYLE_CHARS_UPPER_LETTER(0), Numbering is put in upper case letters. ("A, B, C, D)
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
;Counting Type Constants: $LOW_FOOTNOTE_COUNT_PER_PAGE(0), Restarts the numbering of footnotes at the top of each page. This
;								option is only available if End of Doc is set to False.
;							$LOW_FOOTNOTE_COUNT_PER_CHAP(1), Restarts the numbering of footnotes at the beginning of each
;								chapter.
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

	;0 Based -- Minus 1
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sEnd                - [optional] a string value. Default is Null. The text to display at the end of a
;				   +						Footnote before it continues on the next page.
;                  $sBegin              - [optional] a string value. Default is Null. The text to display at the beginning of a
;				   +						Footnote that has continued on the next page.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sEnd not a String.
;				   @Error 1 @Extended 3 Return 0 = $sBegin not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2
;				   |								1 = Error setting $sEnd
;				   |								2 = Error setting $sBegin
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sParagraph          - [optional] a string value. Default is Null. The Footnote Text Paragraph Style.
;                  $sPage               - [optional] a string value. Default is Null. The Page Style to use for the Footnote
;				   +						pages. Only valid if the footnotes are set to End of Document, instead of per page.
;                  $sTextArea           - [optional] a string value. Default is Null. The Character Style to use for the Footnote
;				   +						anchor in the document text.
;                  $sFootnoteArea       - [optional] a string value. Default is Null. The Character Style to use for the Footnote
;				   +						number in the footnote text.
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
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for
;				   +								the following values: 1, 2, 4, 8
;				   |								1 = Error setting $sParagraph
;				   |								2 = Error setting $sPage
;				   |								4 = Error setting $sTextArea
;				   |								8 = Error setting $sFootnoteArea
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Footnotes Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully searched for Footnotes, none contained in document.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for Footnotes, Returning Array of Footnote
;				   +										Objects. @Extended set to number found.
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
; Name ..........: _LOWriter_FormatKeyCreate
; Description ...: Create a Format Key.
; Syntax ........: _LOWriter_FormatKeyCreate(Byref $oDoc, $sFormat)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $sFormat             - a string value. The format String to create.
; Return values .:  Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFormat not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to Create or Retrieve the Format key, but failed.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Format Key was successfully created, returning Format Key
;				   +												integer.
;				   @Error 0 @Extended 1 Return Integer = Success. Format Key already existed, returning Format Key integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormatKeyDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormatKeyCreate(ByRef $oDoc, $sFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFormatKey
	Local $tLocale
	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	$iFormatKey = $oFormats.queryKey($sFormat, $tLocale, False)
	If ($iFormatKey > -1) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $iFormatKey) ;Format already existed
	$iFormatKey = $oFormats.addNew($sFormat, $tLocale)
	If ($iFormatKey > -1) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $iFormatKey) ;Format created

	Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ;Failed to create or retrieve Format
EndFunc   ;==>_LOWriter_FormatKeyCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormatKeyDelete
; Description ...: Delete a User-Created Format Key from a Document.
; Syntax ........: _LOWriter_FormatKeyDelete(Byref $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $iFormatKey          - an integer value. The User-Created format Key to delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = Format Key called in $iFormatKey not found in Document.
;				   @Error 1 @Extended 4 Return 0 = Format Key called in $iFormatKey not User-Created.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to delete key, but Key is still found in Document.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Format Key was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormatKeyList, _LOWriter_FormatKeyCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormatKeyDelete(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tLocale
	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_FormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ;Key not found.
	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	If ($oFormats.getbykey($iFormatKey).UserDefined() = False) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0) ;Key not User Created.

	$oFormats.removeByKey($iFormatKey)

	Return (_LOWriter_FormatKeyExists($oDoc, $iFormatKey) = False) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>_LOWriter_FormatKeyDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormatKeyExists
; Description ...:Check if a Document contains a certain Format Key.
; Syntax ........: _LOWriter_FormatKeyExists(Byref $oDoc, $iFormatKey, Const $iFormatType)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $iFormatKey          - an integer value. The Format Key to look for.
;                  $iFormatType         - [optional] an integer value. Default is $LOW_FORMAT_KEYS_ALL. The Formatk Key type to
;				   +					search in. Values can be BitOr's together. See Constants.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iFormatType not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to obtain Array of Date/Time Formats.
;				   --Success--
;				   @Error 0 @Extended 0 Return True = Success. Format Key exists in document.
;				   @Error 0 @Extended 1 Return False = Success. Format Key does not exist in document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Format Key Type Constants: $LOW_FORMAT_KEYS_ALL(0), All number formats.
;							$LOW_FORMAT_KEYS_DEFINED(1), Only user-defined number formats.
;							$LOW_FORMAT_KEYS_DATE(2), Date formats.
;							$LOW_FORMAT_KEYS_TIME(4), Time formats.
;							$LOW_FORMAT_KEYS_DATE_TIME(6), Number formats which contain date and time.
;							$LOW_FORMAT_KEYS_CURRENCY(8), Currency formats.
;							$LOW_FORMAT_KEYS_NUMBER(16), Decimal number formats.
;							$LOW_FORMAT_KEYS_SCIENTIFIC(32), Scientific number formats.
;							$LOW_FORMAT_KEYS_FRACTION(64), Number formats for fractions.
;							$LOW_FORMAT_KEYS_PERCENT(128), Percentage number formats.
;							$LOW_FORMAT_KEYS_TEXT(256), Text number formats.
;							$LOW_FORMAT_KEYS_LOGICAL(1024), Boolean number formats.
;							$LOW_FORMAT_KEYS_UNDEFINED(2048), Is used as a return value if no format exists.
;							$LOW_FORMAT_KEYS_EMPTY(4096),
;							$LOW_FORMAT_KEYS_DURATION(8196),
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormatKeyExists(ByRef $oDoc, $iFormatKey, $iFormatType = $LOW_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys[0]
	Local $tLocale

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iFormatType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	$aiFormatKeys = $oFormats.queryKeys($iFormatType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	For $i = 0 To UBound($aiFormatKeys) - 1
		If ($aiFormatKeys[$i] = $iFormatKey) Then Return SetError($__LOW_STATUS_SUCCESS, 0, True) ;Doc does contain format Key
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
	Next

	Return SetError($__LOW_STATUS_SUCCESS, 1, False) ;Doc does not contain format Key
EndFunc   ;==>_LOWriter_FormatKeyExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormatKeyGetString
; Description ...:  Retrieve a Format Key String.
; Syntax ........: _LOWriter_FormatKeyGetString(Byref $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $iFormatKey          - an integer value. The Format Key to retrieve the string for.
; Return values .:Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iFormatKey not found in Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve requested Format Key Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returning Format Key's Format String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormatKeyList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormatKeyGetString(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormatKey

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_FormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oFormatKey = $oDoc.getNumberFormats().getByKey($iFormatKey)
	If Not IsObj($oFormatKey) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0) ;Key not found.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oFormatKey.FormatString())
EndFunc   ;==>_LOWriter_FormatKeyGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormatKeyList
; Description ...: Retrieve an Array of Date/Time Format Keys.
; Syntax ........: _LOWriter_FormatKeyList(Byref $oDoc[, $bIsUser = False[, $bUserOnly = False[, $iFormatKeyType = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $bIsUser             - [optional] a boolean value. Default is False. If True, Adds a third column to the
;				   +						return Array with a boolean, whether each Key is user-created or not.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True, only user-created Format Keys
;				   +						are returned.
;                  $iFormatKeyType      - [optional] an integer value. Default is $LOW_FORMAT_KEYS_ALL. The Formatk Key type to
;				   +					retrieve a list for. Values can be BitOr's together. See Constants.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsUser not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bUserOnly not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iFormatKeyType not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve NumberFormats Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to obtain Array of Format Keys.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning a 2 or three column Array, depending on current
;				   +										$bIsUser setting. Column One (Array[0][0]) will contain the Format
;				   +										Key integer, Column two (Array[0][1]) will contain the Format Key
;				   +										String, If $bIsUser is set to True, Column Three (Array[0][2]) will
;				   +										contain a Boolean, True if the Format Key is User created, else
;				   +										false. @Extended is set to the number of Keys returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
;Format Key Type Constants: $LOW_FORMAT_KEYS_ALL(0), All number formats.
;							$LOW_FORMAT_KEYS_DEFINED(1), Only user-defined number formats.
;							$LOW_FORMAT_KEYS_DATE(2), Date formats.
;							$LOW_FORMAT_KEYS_TIME(4), Time formats.
;							$LOW_FORMAT_KEYS_DATE_TIME(6), Number formats which contain date and time.
;							$LOW_FORMAT_KEYS_CURRENCY(8), Currency formats.
;							$LOW_FORMAT_KEYS_NUMBER(16), Decimal number formats.
;							$LOW_FORMAT_KEYS_SCIENTIFIC(32), Scientific number formats.
;							$LOW_FORMAT_KEYS_FRACTION(64), Number formats for fractions.
;							$LOW_FORMAT_KEYS_PERCENT(128), Percentage number formats.
;							$LOW_FORMAT_KEYS_TEXT(256), Text number formats.
;							$LOW_FORMAT_KEYS_LOGICAL(1024), Boolean number formats.
;							$LOW_FORMAT_KEYS_UNDEFINED(2048), Is used as a return value if no format exists.
;							$LOW_FORMAT_KEYS_EMPTY(4096),
;							$LOW_FORMAT_KEYS_DURATION(8196),
; Related .......: _LOWriter_FormatKeyDelete, _LOWriter_FormatKeyGetString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormatKeyList(ByRef $oDoc, $bIsUser = False, $bUserOnly = False, $iFormatKeyType = $LOW_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys
	Local $avFormats[0][3]
	Local $tLocale
	Local $iColumns = 3, $iCount = 0

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bIsUser) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$iColumns = ($bIsUser = True) ? $iColumns : 2

	If Not IsInt($iFormatKeyType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	$aiFormatKeys = $oFormats.queryKeys($iFormatKeyType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	ReDim $avFormats[UBound($aiFormatKeys)][$iColumns]

	For $i = 0 To UBound($aiFormatKeys) - 1

		If ($bUserOnly = True) Then
			If ($oFormats.getbykey($aiFormatKeys[$i]).UserDefined() = True) Then
				$avFormats[$iCount][0] = $aiFormatKeys[$i]
				$avFormats[$iCount][1] = $oFormats.getbykey($aiFormatKeys[$i]).FormatString()
				If ($bIsUser = True) Then $avFormats[$iCount][2] = $oFormats.getbykey($aiFormatKeys[$i]).UserDefined()
				$iCount += 1
			EndIf
		Else
			$avFormats[$i][0] = $aiFormatKeys[$i]
			$avFormats[$i][1] = $oFormats.getbykey($aiFormatKeys[$i]).FormatString()
			If ($bIsUser = True) Then $avFormats[$i][2] = $oFormats.getbykey($aiFormatKeys[$i]).UserDefined()
		EndIf
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
	Next

	If ($bUserOnly = True) Then ReDim $avFormats[$iCount][$iColumns]

	Return SetError($__LOW_STATUS_SUCCESS, UBound($avFormats), $avFormats)
EndFunc   ;==>_LOWriter_FormatKeyList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_SearchDescriptorCreate
; Description ...: Create a Search Descriptor for searching a document.
; Syntax ........: _LOWriter_SearchDescriptorCreate(Byref $oDoc[, $bBackwards = False[, $bMatchCase = False[, $bWholeWord = False[, $bRegExp = False[, $bStyles = False[, $bSearchPropValues = False]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
;                  $bBackwards          - [optional] a boolean value. Default is False. If True, the document is searched
;				   +						backwards.
;                  $bMatchCase          - [optional] a boolean value. Default is False. If True, the case of the letters is
;				   +						important for the Search.
;                  $bWholeWord          - [optional] a boolean value. Default is False. If True, only complete words will be
;				   +						found.
;                  $bRegExp             - [optional] a boolean value. Default is False. If True, the search string is evaluated
;				   +						as a regular expression.
;                  $bStyles             - [optional] a boolean value. Default is False. If True, the string is considered a
;				   +						Paragraph Style name, and the search will return any paragraph utilizing the
;				   +						specified name, EXCEPT if you input Format properties to search for, then setting
;				   +						this to True causes the search to search both for direct formatting matching those
;				   +						properties and also Paragraph/Character styles that contain matching properties.
;                  $bSearchPropValues   - [optional] a boolean value. Default is False. If True, any formatting properties
;				   +						searched for are matched based on their value, else if false, the search only looks
;				   +						for their existence. See Remarks.
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
; Remarks .......: $bSearchPropValues is equivalent to the difference in selecting "Format" options in Libre Office's search
;						box and "Attributes". Setting $bSearchPropValues to True, means that the search will look for matches
;						using the specified property AND having the specified value, such as Character Weight, Bold, only
;						matches that have Character weight of Bold will be returned, whereas if $bSearchPropValues is set to
;						false, the search only looks for matches that have the specified property, regardless of its value. Such
;						as Character weight, would match Bold, Semi-Bold, etc. From my understanding, the search is based on
;						anything directly formatted unless $bStyles is also true.
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
; Parameters ....: $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from
;				   +						_LOWriter_SearchDescriptorCreate function.
;                  $bBackwards          - [optional] a boolean value. Default is False. If True, the document is searched
;				   +						backwards.
;                  $bMatchCase          - [optional] a boolean value. Default is False. If True, the case of the letters is
;				   +						important for the Search.
;                  $bWholeWord          - [optional] a boolean value. Default is False. If True, only complete words will be
;				   +						found.
;                  $bRegExp             - [optional] a boolean value. Default is False. If True, the search string is evaluated
;				   +						as a regular expression. Cannot be set to True if Similarity Search is set to True.
;                  $bStyles             - [optional] a boolean value. Default is False. If True, the string is considered a
;				   +						Paragraph Style name, and the search will return any paragraph utilizing the
;				   +						specified name, EXCEPT if you input Format properties to search for, then setting
;				   +						this to True causes the search to search both for direct formatting matching those
;				   +						properties and also Paragraph/Character styles that contain matching properties.
;                  $bSearchPropValues   - [optional] a boolean value. Default is False. If True, any formatting properties
;				   +						searched for are matched based on their value, else if false, the search only looks
;				   +						for their existence. See Remarks.
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
;;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bSearchPropValues is equivalent to the difference in selecting "Format" options in Libre Office's search
;						box and "Attributes". Setting $bSearchPropValues to True, means that the search will look for matches
;						using the specified property AND having the specified value, such as Character Weight, Bold, only
;						matches that have Character weight of Bold will be returned, whereas if $bSearchPropValues is set to
;						false, the search only looks for matches that have the specified property, regardless of its value. Such
;						as Character weight, would match Bold, Semi-Bold, etc. From my understanding, the search is based on
;						anything directly formatted unless $bStyles is also true.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;						get the current settings.
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
; Parameters ....: $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from
;				   +						_LOWriter_SearchDescriptorCreate function.
;                  $bSimilarity         - [optional] a boolean value. Default is Null. If True, a "similarity search" is
;				   +						performed.
;                  $bCombine            - [optional] a boolean value. Default is Null. If True, all similarity rules ($iRemove,
;				   +						$iAdd, and $iExchange) are applied together.
;                  $iRemove             - [optional] an integer value. Default is Null. Specifies the number of characters that
;				   +						may be ignored to match the search pattern.
;                  $iAdd                - [optional] an integer value. Default is Null. Specifies the number of characters that
;				   +						must be added to match the search pattern.
;                  $iExchange           - [optional] an integer value. Default is Null. Specifies the number of characters that
;				   +						must be replaced to match the search pattern.
; Return values .:  Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript Object not a Search Descriptor Object.
;				   @Error 1 @Extended 3 Return 0 = $bSimilarity not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bCombine not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iRemove, $iAdd, or $iExchange set to a value, but $bSimilarity not set to
;				   +									True.
;				   @Error 1 @Extended 6 Return 0 = $iRemove not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iAdd not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iExchange not an Integer.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = $bSimilarity is set to True while Regular Expression Search is also set to
;				   +									True.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Returns 1 after directly modifying Search Descriptor Object.
;;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current
;				   +								settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;						get the current settings.
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
; Parameters ....: $oDoc                - [in/out] an object.  A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
; Return values .: Success: 2D Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Shapes Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning 2D Array containing a list of Shape names contained
;				   +									contained in a document, the first column ($aArray[0][0] contains the
;				   +									shape name, the second column ($aArray[0][1] contains the shape's
;				   +									Implementation name. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Implementation name identifies what type of shape object it is, as there can be multiple things counted
;						as "Shapes", such as Text Frames etc.
;						I have found the three Implementation names being returned, SwXTextFrame, indicating the shape
;						is actually a Text Frame, SwXShape, is a regular shape such as a line, circle etc. And
;						"SwXTextGraphicObject", which is an image / picture. There may be other return types I haven't found
;						yet. Images inserted into the document are also listed as TextFrames in the shapes category. There isn't
;						and easy way to differentiate between them yet, see _LOWriter_FramesGetNames, to search for Frames in
;						the shapes category.
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
