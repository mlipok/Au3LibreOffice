#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;~ #Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Listing and querying available L.O. Writer Fonts.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_FontDescCreate
; _LOWriter_FontDescEdit
; _LOWriter_FontExists
; _LOWriter_FontsGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FontDescCreate
; Description ...: Create a Font Descriptor Map.
; Syntax ........: _LOWriter_FontDescCreate([$sFontName = ""[, $iWeight = $LOW_WEIGHT_DONT_KNOW[, $iSlant = $LOW_POSTURE_DONTKNOW[, $nSize = 0[, $iColor = $LOW_COLOR_OFF[, $iUnderlineStyle = $LOW_UNDERLINE_DONT_KNOW[, $iUnderlineColor = $LOW_COLOR_OFF[, $iStrikelineStyle = $LOW_STRIKEOUT_DONT_KNOW[, $bIndividualWords = False[, $iRelief = $LOW_RELIEF_NONE]]]]]]]]]])
; Parameters ....: $sFontName           - [optional] a string value. Default is "". The Font name.
;                  $iWeight             - [optional] an integer value (0-200). Default is $LOW_WEIGHT_DONT_KNOW. The Font weight. See Constants $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iSlant              - [optional] an integer value (0-5). Default is $LOW_POSTURE_DONTKNOW. The Font italic setting. See Constants $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $nSize               - [optional] a general number value. Default is 0. The Font size.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is $LOW_COLOR_OFF. The Font Color in Long Integer format, can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for Auto color.
;                  $iUnderlineStyle     - [optional] an integer value (0-18). Default is $LOW_UNDERLINE_DONT_KNOW. The Font underline Style. See Constants $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iUnderlineColor     - [optional] an integer value (-1-16777215). Default is $LOW_COLOR_OFF. The Font Underline color in Long Integer format, can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for Auto color.
;                  $iStrikelineStyle    - [optional] an integer value (0-6). Default is $LOW_STRIKEOUT_DONT_KNOW. The Strikeout line style. See Constants $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bIndividualWords    - [optional] a boolean value. Default is False. If True, only individual words are underlined.
;                  $iRelief             - [optional] an integer value (0-2). Default is $LOW_RELIEF_NONE. The Font relief style. See Constants $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 2 Return 0 = Font called in $sFontName not found.
;                  @Error 1 @Extended 3 Return 0 = $iWeight not an Integer, less than 0 or greater than 200. See Constants $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iSlant not an Integer, less than 0 or greater than 5. See Constants $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $nSize not a number.
;                  @Error 1 @Extended 6 Return 0 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 7 Return 0 = $iUnderlineStyle not an Integer, less than 0 or greater than 18. See Constants $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iUnderlineColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 9 Return 0 = $iStrikelineStyle not an Integer, less than 0 or greater than 6. See Constants $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $bIndividualWords not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $iRelief not an Integer, less than 0 or greater than 2. See Constants $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Returning the created Map Font Descriptor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FontDescCreate($sFontName = "", $iWeight = $LOW_WEIGHT_DONT_KNOW, $iSlant = $LOW_POSTURE_DONTKNOW, $nSize = 0, $iColor = $LOW_COLOR_OFF, $iUnderlineStyle = $LOW_UNDERLINE_DONT_KNOW, $iUnderlineColor = $LOW_COLOR_OFF, $iStrikelineStyle = $LOW_STRIKEOUT_DONT_KNOW, $bIndividualWords = False, $iRelief = $LOW_RELIEF_NONE)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $mFontDesc[]

	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not _LOWriter_FontExists($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_IntIsBetween($iWeight, $LOW_WEIGHT_DONT_KNOW, $LOW_WEIGHT_BLACK) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LOWriter_IntIsBetween($iSlant, $LOW_POSTURE_NONE, $LOW_POSTURE_REV_ITALIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsNumber($nSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not __LOWriter_IntIsBetween($iColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not __LOWriter_IntIsBetween($iUnderlineStyle, $LOW_UNDERLINE_NONE, $LOW_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not __LOWriter_IntIsBetween($iUnderlineColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If Not __LOWriter_IntIsBetween($iStrikelineStyle, $LOW_STRIKEOUT_NONE, $LOW_STRIKEOUT_X) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If Not IsBool($bIndividualWords) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
	If Not __LOWriter_IntIsBetween($iRelief, $LOW_RELIEF_NONE, $LOW_RELIEF_ENGRAVED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

	$mFontDesc.CharFontName = $sFontName
	$mFontDesc.CharWeight = $iWeight
	$mFontDesc.CharPosture = $iSlant
	$mFontDesc.CharHeight = $nSize
	$mFontDesc.CharColor = $iColor
	$mFontDesc.CharUnderline = $iUnderlineStyle
	$mFontDesc.CharUnderlineColor = $iUnderlineColor
	$mFontDesc.CharStrikeout = $iStrikelineStyle
	$mFontDesc.CharWordMode = $bIndividualWords
	$mFontDesc.CharRelief = $iRelief

	Return SetError($__LO_STATUS_SUCCESS, 0, $mFontDesc)
EndFunc   ;==>_LOWriter_FontDescCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FontDescEdit
; Description ...: Set or Retrieve Font Descriptor settings.
; Syntax ........: _LOWriter_FontDescEdit(ByRef $mFontDesc[, $sFontName = Null[, $iWeight = Null[, $iSlant = Null[, $nSize = Null[, $iColor = Null[, $iUnderlineStyle = Null[, $iUnderlineColor = Null[, $iStrikelineStyle = Null[, $bIndividualWords = Null[, $iRelief = Null]]]]]]]]]])
; Parameters ....: $mFontDesc           - [in/out] a map. A Font descriptor Map as returned from a _LOWriter_FontDescCreate, or control property return function.
;                  $sFontName           - [optional] a string value. Default is Null. The Font name.
;                  $iWeight             - [optional] an integer value (0-200). Default is Null. The Font weight. See Constants $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iSlant              - [optional] an integer value (0-5). Default is Null. The Font italic setting. See Constants $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $nSize               - [optional] a general number value. Default is Null. The Font size.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is Null. The Font Color in Long Integer format, can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for Auto color.
;                  $iUnderlineStyle     - [optional] an integer value (0-18). Default is Null. The Font underline Style. See Constants $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iUnderlineColor     - [optional] an integer value (-1-16777215). Default is Null.
;                  $iStrikelineStyle    - [optional] an integer value (0-6). Default is Null. The Strikeout line style. See Constants $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bIndividualWords    - [optional] a boolean value. Default is Null. If True, only individual words are underlined.
;                  $iRelief             - [optional] an integer value (0-2). Default is Null. The Font relief style. See Constants $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $mFontDesc not a Map.
;                  @Error 1 @Extended 2 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 3 Return 0 = Font called in $sFontName not found.
;                  @Error 1 @Extended 4 Return 0 = $iWeight not an Integer, less than 0 or greater than 200. See Constants $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iSlant not an Integer, less than 0 or greater than 5. See Constants $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $nSize not a number.
;                  @Error 1 @Extended 7 Return 0 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 8 Return 0 = $iUnderlineStyle not an Integer, less than 0 or greater than 18. See Constants $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $iUnderlineColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 10 Return 0 = $iStrikelineStyle not an Integer, less than 0 or greater than 6. See Constants $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = $bIndividualWords not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iRelief not an Integer, less than 0 or greater than 2. See Constants $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 10 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FontDescEdit(ByRef $mFontDesc, $sFontName = Null, $iWeight = Null, $iSlant = Null, $nSize = Null, $iColor = Null, $iUnderlineStyle = Null, $iUnderlineColor = Null, $iStrikelineStyle = Null, $bIndividualWords = Null, $iRelief = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avFont[10]

	If Not IsMap($mFontDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sFontName, $iWeight, $iSlant, $nSize, $iColor, $iUnderlineStyle, $iUnderlineColor, $iStrikelineStyle, $bIndividualWords, $iRelief) Then
		__LOWriter_ArrayFill($avFont, $mFontDesc.CharFontName, $mFontDesc.CharWeight, $mFontDesc.CharPosture, $mFontDesc.CharHeight, $mFontDesc.CharColor, $mFontDesc.CharUnderline, _
				$mFontDesc.CharUnderlineColor, $mFontDesc.CharStrikeout, $mFontDesc.CharWordMode, $mFontDesc.CharRelief)
		Return SetError($__LO_STATUS_SUCCESS, 1, $avFont)
	EndIf

	If ($sFontName <> Null) Then
		If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		If Not _LOWriter_FontExists($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$mFontDesc.CharFontName = $sFontName
	EndIf

	If ($iWeight <> Null) Then
		If Not __LOWriter_IntIsBetween($iWeight, $LOW_WEIGHT_DONT_KNOW, $LOW_WEIGHT_BLACK) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$mFontDesc.CharWeight = $iWeight
	EndIf

	If ($iSlant <> Null) Then
		If Not __LOWriter_IntIsBetween($iSlant, $LOW_POSTURE_NONE, $LOW_POSTURE_REV_ITALIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$mFontDesc.CharPosture = $iSlant
	EndIf

	If ($nSize <> Null) Then
		If Not IsNumber($nSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$mFontDesc.CharHeight = $nSize
	EndIf

	If ($iColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$mFontDesc.CharColor = $iColor
	EndIf

	If ($iUnderlineStyle <> Null) Then
		If Not __LOWriter_IntIsBetween($iUnderlineStyle, $LOW_UNDERLINE_NONE, $LOW_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$mFontDesc.CharUnderline = $iUnderlineStyle
	EndIf

	If ($iUnderlineColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iUnderlineColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		$mFontDesc.CharUnderlineColor = $iUnderlineColor
	EndIf

	If ($iStrikelineStyle <> Null) Then
		If Not __LOWriter_IntIsBetween($iStrikelineStyle, $LOW_STRIKEOUT_NONE, $LOW_STRIKEOUT_X) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		$mFontDesc.CharStrikeout = $iStrikelineStyle
	EndIf

	If ($bIndividualWords <> Null) Then
		If Not IsBool($bIndividualWords) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
		$mFontDesc.CharWordMode = $bIndividualWords
	EndIf

	If ($iRelief <> Null) Then
		If Not __LOWriter_IntIsBetween($iRelief, $LOW_RELIEF_NONE, $LOW_RELIEF_ENGRAVED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)
		$mFontDesc.CharRelief = $iRelief
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FontDescEdit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FontExists
; Description ...: Tests whether a Document has a specific font available by name.
; Syntax ........: _LOWriter_FontExists($sFontName)
; Parameters ....: $sFontName           - a string value. The Font name to search for.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFontName not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  @Error 2 @Extended 3 Return 0 = Failed to create a Property Struct.
;                  @Error 2 @Extended 4 Return 0 = Failed to create a new Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Font list.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the Font is available, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function may cause a processor usage spike for a moment or two. If you wish to eliminate this, comment out the current sleep function and place a sleep(10) in its place.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FontExists($sFontName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts
	Local Const $iURLFrameCreate = 8 ;frame will be created if not found
	Local $oServiceManager, $oDesktop, $oDoc
	Local $atProperties[1]

	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$atProperties[0] = __LOWriter_SetPropertyValue("Hidden", True)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	$oDoc = $oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", $iURLFrameCreate, $atProperties)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	For $i = 0 To UBound($atFonts) - 1
		If $atFonts[$i].Name() = $sFontName Then
			$oDoc.Close(True)
			Return SetError($__LO_STATUS_SUCCESS, 0, True)
		EndIf

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOWriter_FontExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FontsGetNames
; Description ...: Retrieve an array of currently available fonts.
; Syntax ........: _LOWriter_FontsGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Font list.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returns a 4 Column Array, @extended is set to the number of results. See remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Many fonts will be listed multiple times, this is because of the varying settings for them, such as bold, Italic, etc.
;                  Style Name is really a repeat of weight(Bold) and Slant (Italic) settings, but is included for easier processing if required.
;                  From personal tests, Slant only returns 0 or 2. This function may cause a processor usage spike for a moment or two.
;                  The returned array will be as follows:
;                  The first column (Array[1][0]) contains the Font Name.
;                  The Second column (Array [1][1] contains the style name (Such as Bold Italic etc.)
;                  The third column (Array[1][2]) contains the Font weight (Bold) See Constants, $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3;
;                  The fourth column (Array[1][3]) contains the font slant (Italic) See constants, $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FontsGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts
	Local $asFonts[0][4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asFonts[UBound($atFonts)][4]

	For $i = 0 To UBound($atFonts) - 1
		$asFonts[$i][0] = $atFonts[$i].Name()
		$asFonts[$i][1] = $atFonts[$i].StyleName()
		$asFonts[$i][2] = $atFonts[$i].Weight
		$asFonts[$i][3] = $atFonts[$i].Slant() ; only 0 or 2?
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($atFonts), $asFonts)
EndFunc   ;==>_LOWriter_FontsGetNames
