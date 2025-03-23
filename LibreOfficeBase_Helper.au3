#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;~ #Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Base
#include "LibreOfficeBase_Constants.au3"
#include "LibreOfficeBase_Internal.au3"

; Other includes for Base

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Functions used for creating, modifying and retrieving data for use in various functions in LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_ComError_UserFunction
; _LOBase_DateStructCreate
; _LOBase_DateStructModify
; _LOBase_PathConvert
; _LOBase_VersionGet
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LOBase_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
; Parameters ....: $vUserFunction       - [optional] a Function or Keyword. Default value is Default. Accepts a Function, or the Keyword Default and Null. If set to a User function, the function may have up to 5 required parameters.
;                  $vParam1             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam2             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam3             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam4             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam5             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
; Return values .: Success: 1 or UserFunction.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $vUserFunction Not a Function, or Default keyword, or Null Keyword.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Successfully set the UserFunction.
;                  @Error 0 @Extended 0 Return 2 = Successfully cleared the set UserFunction.
;                  @Error 0 @Extended 0 Return Function = Returning the set UserFunction.
; Author ........: mLipok
; Modified ......: donnyh13 - Added a clear UserFunction without error option. Also added parameters option.
; Remarks .......: The first parameter passed to the User function will always be the COM Error object. See below.
;                  Every COM Error will be passed to that function. The user can then read the following properties. (As Found in the COM Reference section in AutoIt Help File.) Using the first parameter in the UserFunction.
;                  For Example MyFunc($oMyError)
;                    $oMyError.number The Windows HRESULT value from a COM call
;                    $oMyError.windescription The FormatWinError() text derived from .number
;                    $oMyError.source Name of the Object generating the error (contents from ExcepInfo.source)
;                    $oMyError.description Source Object's description of the error (contents from ExcepInfo.description)
;                    $oMyError.helpfile Source Object's help file for the error (contents from ExcepInfo.helpfile)
;                    $oMyError.helpcontext Source Object's help file context id number (contents from ExcepInfo.helpcontext)
;                    $oMyError.lastdllerror The number returned from GetLastError()
;                    $oMyError.scriptline The script line on which the error was generated
;                    NOTE: Not all properties will necessarily contain data, some will be blank.
;                  If MsgBox or ConsoleWrite functions are passed to this function, the error details will be displayed using that function automatically.
;                  If called with Default keyword, the current UserFunction, if set, will be returned.
;                  If called with Null keyword, the currently set UserFunction is cleared and only the internal ComErrorHandler will be called for COM Errors.
;                  The stored UserFunction (besides MsgBox and ConsoleWrite) will be called as follows: UserFunc($oComError,$vParam1,$vParam2,$vParam3,$vParam4,$vParam5)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
	#forceref $vParam1, $vParam2, $vParam3, $vParam4, $vParam5
	; If user does not set a function, UDF must use internal function to avoid AutoItError.
	Local Static $vUserFunction_Static = Default
	Local $avUserFuncWParams[@NumParams]

	If $vUserFunction = Default Then
		; just return stored static User Function variable
		Return SetError($__LO_STATUS_SUCCESS, 0, $vUserFunction_Static)
	ElseIf IsFunc($vUserFunction) Then
		; If User called Parameters, then add to array.
		If @NumParams > 1 Then
			$avUserFuncWParams[0] = $vUserFunction
			For $i = 1 To @NumParams - 1
				$avUserFuncWParams[$i] = Eval("vParam" & $i)
				; set static variable
			Next
			$vUserFunction_Static = $avUserFuncWParams
		Else
			$vUserFunction_Static = $vUserFunction
		EndIf
		Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	ElseIf $vUserFunction = Null Then
		; Clear User Function.
		$vUserFunction_Static = Default
		Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	Else
		; return error as an incorrect parameter was passed to this function
		Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOBase_ComError_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DateStructCreate
; Description ...: Create a Date Structure for inserting a Date into certain other functions.
; Syntax ........: _LOBase_DateStructCreate([$iYear = Null[, $iMonth = Null[, $iDay = Null[, $iHours = Null[, $iMinutes = Null[, $iSeconds = Null[, $iNanoSeconds = Null[, $bIsUTC = Null]]]]]]]])
; Parameters ....: $iYear               - [optional] an integer value. Default is Null. The Year, in 4 digit integer format.
;                  $iMonth              - [optional] an integer value (0-12). Default is Null. The Month, in 2 digit integer format. Set to 0 for Void date.
;                  $iDay                - [optional] an integer value (0-31). Default is Null. The Day, in 2 digit integer format. Set to 0 for Void date.
;                  $iHours              - [optional] an integer value (0-23). Default is Null. The Hour, in 2 digit integer format.
;                  $iMinutes            - [optional] an integer value (0-59). Default is Null. Minutes, in 2 digit integer format.
;                  $iSeconds            - [optional] an integer value (0-59). Default is Null. Seconds, in 2 digit integer format.
;                  $iNanoSeconds        - [optional] an integer value (0-999,999,999). Default is Null. Nano-Second, in integer format. Min 0, Max 999,999,999.
;                  $bIsUTC              - [optional] a boolean value. Default is Null. If true: time zone is UTC Else false: unknown time zone. Libre Office version 4.1 and up.
; Return values .: Success: Structure.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iYear not an Integer.
;                  @Error 1 @Extended 2 Return 0 = $iYear not 4 digits long.
;                  @Error 1 @Extended 3 Return 0 = $iMonth not an Integer, less than 0, or greater than 12.
;                  @Error 1 @Extended 4 Return 0 = $iDay not an Integer, less than 0, or greater than 31.
;                  @Error 1 @Extended 5 Return 0 = $iHours not an Integer, less than 0, or greater than 23.
;                  @Error 1 @Extended 6 Return 0 = $iMinutes not an Integer, less than 0, or greater than 59.
;                  @Error 1 @Extended 7 Return 0 = $iSeconds not an Integer, less than 0, or greater than 59.
;                  @Error 1 @Extended 8 Return 0 = $iNanoSeconds not an Integer, less than 0, or greater than 999999999.
;                  @Error 1 @Extended 9 Return 0 = $bIsUTC not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.util.DateTime" Object.
;                  --Version Related Errors--
;                  @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.1.
;                  --Success--
;                  @Error 0 @Extended 0 Return Structure = Success. Successfully created the Date/Time Structure, Returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling a value with Null keyword will auto fill the value with the current value, such as current hour, etc.
; Related .......: _LOBase_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DateStructCreate($iYear = Null, $iMonth = Null, $iDay = Null, $iHours = Null, $iMinutes = Null, $iSeconds = Null, $iNanoSeconds = Null, $bIsUTC = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tDateStruct

	$tDateStruct = __LOBase_CreateStruct("com.sun.star.util.DateTime")
	If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($iYear <> Null) Then
		If Not IsInt($iYear) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
		If Not (StringLen($iYear) = 4) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$tDateStruct.Year = $iYear
	Else
		$tDateStruct.Year = @YEAR
	EndIf

	If ($iMonth <> Null) Then
		If Not __LOBase_IntIsBetween($iMonth, 0, 12) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$tDateStruct.Month = $iMonth
	Else
		$tDateStruct.Month = @MON
	EndIf

	If ($iDay <> Null) Then
		If Not __LOBase_IntIsBetween($iDay, 0, 31) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$tDateStruct.Day = $iDay
	Else
		$tDateStruct.Day = @MDAY
	EndIf

	If ($iHours <> Null) Then
		If Not __LOBase_IntIsBetween($iHours, 0, 23) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$tDateStruct.Hours = $iHours
	Else
		$tDateStruct.Hours = @HOUR
	EndIf

	If ($iMinutes <> Null) Then
		If Not __LOBase_IntIsBetween($iMinutes, 0, 59) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$tDateStruct.Minutes = $iMinutes
	Else
		$tDateStruct.Minutes = @MIN
	EndIf

	If ($iSeconds <> Null) Then
		If Not __LOBase_IntIsBetween($iSeconds, 0, 59) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$tDateStruct.Seconds = $iSeconds
	Else
		$tDateStruct.Seconds = @SEC
	EndIf

	If ($iNanoSeconds <> Null) Then
		If Not __LOBase_IntIsBetween($iNanoSeconds, 0, 999999999) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$tDateStruct.NanoSeconds = $iNanoSeconds
	Else
		$tDateStruct.NanoSeconds = 0
	EndIf

	If ($bIsUTC <> Null) Then
		If Not IsBool($bIsUTC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not __LOBase_VersionCheck(4.1) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
		$tDateStruct.IsUTC = $bIsUTC
	Else
		If __LOBase_VersionCheck(4.1) Then $tDateStruct.IsUTC = False
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $tDateStruct)
EndFunc   ;==>_LOBase_DateStructCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DateStructModify
; Description ...: Set or retrieve Date Structure settings.
; Syntax ........: _LOBase_DateStructModify(ByRef $tDateStruct[, $iYear = Null[, $iMonth = Null[, $iDay = Null[, $iHours = Null[, $iMinutes = Null[, $iSeconds = Null[, $iNanoSeconds = Null[, $bIsUTC = Null]]]]]]]])
; Parameters ....: $tDateStruct         - [in/out] a dll struct value. The Date Structure to modify, returned from a _LOBase_DateStructCreate, or setting retrieval function. Structure will be directly modified.
;                  $iYear               - [optional] an integer value. Default is Null. The Year, in 4 digit integer format.
;                  $iMonth              - [optional] an integer value (0-12). Default is Null. The Month, in 2 digit integer format. Set to 0 for Void date.
;                  $iDay                - [optional] an integer value (0-31). Default is Null. The Day, in 2 digit integer format. Set to 0 for Void date.
;                  $iHours              - [optional] an integer value (0-23). Default is Null. The Hour, in 2 digit integer format.
;                  $iMinutes            - [optional] an integer value (0-59). Default is Null. Minutes, in 2 digit integer format.
;                  $iSeconds            - [optional] an integer value (0-59). Default is Null. Seconds, in 2 digit integer format.
;                  $iNanoSeconds        - [optional] an integer value (0-999,999,999). Default is Null. Nano-Second, in integer format.
;                  $bIsUTC              - [optional] a boolean value. Default is Null. If true: time zone is UTC Else false: unknown time zone. Libre Office version 4.1 and up.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $tDateStruct not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iYear not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iYear not 4 digits long.
;                  @Error 1 @Extended 4 Return 0 = $iMonth not an Integer, less than 0, or greater than 12.
;                  @Error 1 @Extended 5 Return 0 = $iDay not an Integer, less than 0, or greater than 31.
;                  @Error 1 @Extended 6 Return 0 = $iHours not an Integer, less than 0, or greater than 23.
;                  @Error 1 @Extended 7 Return 0 = $iMinutes not an Integer, less than 0, or greater than 59.
;                  @Error 1 @Extended 8 Return 0 = $iSeconds not an Integer, less than 0, or greater than 59.
;                  @Error 1 @Extended 9 Return 0 = $iNanoSeconds not an Integer, less than 0, or greater than 999999999.
;                  @Error 1 @Extended 10 Return 0 = $bIsUTC not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iYear
;                  |                               2 = Error setting $iMonth
;                  |                               4 = Error setting $iDay
;                  |                               8 = Error setting $iHours
;                  |                               16 = Error setting $iMinutes
;                  |                               32 = Error setting $iSeconds
;                  |                               64 = Error setting $iNanoSeconds
;                  |                               128 = Error setting $bIsUTC
;                  --Version Related Errors--
;                  @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.1.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 7 or 8 Element Array with values in order of function parameters. If current Libre Office version is less than 4.1, the Array will contain 7 elements, as $bIsUTC will be eliminated.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_DateStructCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DateStructModify(ByRef $tDateStruct, $iYear = Null, $iMonth = Null, $iDay = Null, $iHours = Null, $iMinutes = Null, $iSeconds = Null, $iNanoSeconds = Null, $bIsUTC = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avMod[7]

	If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOBase_VarsAreNull($iYear, $iMonth, $iDay, $iHours, $iMinutes, $iSeconds, $iNanoSeconds, $bIsUTC) Then
		If __LOBase_VersionCheck(4.1) Then
			__LOBase_ArrayFill($avMod, $tDateStruct.Year(), $tDateStruct.Month(), $tDateStruct.Day(), $tDateStruct.Hours(), _
					$tDateStruct.Minutes(), $tDateStruct.Seconds(), $tDateStruct.NanoSeconds(), $tDateStruct.IsUTC())
		Else
			__LOBase_ArrayFill($avMod, $tDateStruct.Year(), $tDateStruct.Month(), $tDateStruct.Day(), $tDateStruct.Hours(), _
					$tDateStruct.Minutes(), $tDateStruct.Seconds(), $tDateStruct.NanoSeconds())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avMod)
	EndIf

	If ($iYear <> Null) Then
		If Not IsInt($iYear) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		If Not (StringLen($iYear) = 4) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$tDateStruct.Year = $iYear
		$iError = ($tDateStruct.Year() = $iYear) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iMonth <> Null) Then
		If Not __LOBase_IntIsBetween($iMonth, 0, 12) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$tDateStruct.Month = $iMonth
		$iError = ($tDateStruct.Month() = $iMonth) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iDay <> Null) Then
		If Not __LOBase_IntIsBetween($iDay, 0, 31) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$tDateStruct.Day = $iDay
		$iError = ($tDateStruct.Day() = $iDay) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iHours <> Null) Then
		If Not __LOBase_IntIsBetween($iHours, 0, 23) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$tDateStruct.Hours = $iHours
		$iError = ($tDateStruct.Hours() = $iHours) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iMinutes <> Null) Then
		If Not __LOBase_IntIsBetween($iMinutes, 0, 59) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$tDateStruct.Minutes = $iMinutes
		$iError = ($tDateStruct.Minutes() = $iMinutes) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iSeconds <> Null) Then
		If Not __LOBase_IntIsBetween($iSeconds, 0, 59) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$tDateStruct.Seconds = $iSeconds
		$iError = ($tDateStruct.Seconds() = $iSeconds) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iNanoSeconds <> Null) Then
		If Not __LOBase_IntIsBetween($iNanoSeconds, 0, 999999999) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		$tDateStruct.NanoSeconds = $iNanoSeconds
		$iError = ($tDateStruct.NanoSeconds() = $iNanoSeconds) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bIsUTC <> Null) Then
		If Not IsBool($bIsUTC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		If Not __LOBase_VersionCheck(4.1) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
		$tDateStruct.IsUTC = $bIsUTC
		$iError = ($tDateStruct.IsUTC() = $bIsUTC) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_DateStructModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_PathConvert
; Description ...: Converts the input path to or from a LibreOffice URL notation path.
; Syntax ........: _LOBase_PathConvert($sFilePath[, $iReturnMode = $LOB_PATHCONV_AUTO_RETURN])
; Parameters ....: $sFilePath           - a string value. Full path to convert in String format.
;                  $iReturnMode         - [optional] an integer value (0-2). Default is $__g_iAutoReturn. The type of path format to return. See Constants, $LOB_PATHCONV_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFilePath is not a string
;                  @Error 1 @Extended 2 Return 0 = $iReturnMode not a Integer, less than 0, or greater than 2, see constants, $LOB_PATHCONV_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Returning converted File Path from Libre Office URL.
;                  @Error 0 @Extended 2 Return String = Returning converted path from File Path to Libre Office URL.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibreOffice URL notation is based on the Internet Standard RFC 1738, which means only [0-9],[a-zA-Z] are allowed in paths, most other characters need to be converted into ISO 8859-1 (ISO Latin) such as is found in internet URL's (spaces become %20). See: StarOfficeTM 6.0 Office SuiteA SunTM ONE Software Offering, Basic Programmer's Guide; Page 74
;                  The user generally should not even need this function, as I have endeavored to convert any URLs to the appropriate computer path format and any input computer paths to a Libre Office URL.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_PathConvert($sFilePath, $iReturnMode = $LOB_PATHCONV_AUTO_RETURN)
	Local Const $STR_STRIPLEADING = 1
	Local $asURLReplace[9][2] = [["%", "%25"], [" ", "%20"], ["\", "/"], [";", "%3B"], ["#", "%23"], ["^", "%5E"], ["{", "%7B"], ["}", "%7D"], ["`", "%60"]]
	Local $iPathSearch, $iFileSearch, $iPartialPCPath, $iPartialFilePath

	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOBase_IntIsBetween($iReturnMode, $LOB_PATHCONV_AUTO_RETURN, $LOB_PATHCONV_PCPATH_RETURN) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sFilePath = StringStripWS($sFilePath, $STR_STRIPLEADING)

	$iPathSearch = StringRegExp($sFilePath, "[A-Z]\:\\") ; Search For a Computer Path, as in C:\ etc.
	$iPartialPCPath = StringInStr($sFilePath, "\") ; Search for partial computer Path containing a backslash.
	$iFileSearch = StringInStr($sFilePath, "file:///", 0, 1, 1, 9) ; Search for a full Libre path, which begins with File:///
	$iPartialFilePath = StringInStr($sFilePath, "/") ; Search For a Partial Libre path containing forward slash

	If ($iReturnMode = $LOB_PATHCONV_AUTO_RETURN) Then

		If ($iPathSearch > 0) Or ($iPartialPCPath > 0) Then ;  if file path contains partial or full PC path, set to convert to Libre URL.
			$iReturnMode = $LOB_PATHCONV_OFFICE_RETURN
		ElseIf ($iFileSearch > 0) Or ($iPartialFilePath > 0) Then ;  if file path contains partial or full Libre URL, set to convert to PC Path.
			$iReturnMode = $LOB_PATHCONV_PCPATH_RETURN
		Else ; If file path contains neither above. convert to Libre URL
			$iReturnMode = $LOB_PATHCONV_OFFICE_RETURN
		EndIf
	EndIf

	Switch $iReturnMode

		Case $LOB_PATHCONV_OFFICE_RETURN
			If $iFileSearch > 0 Then Return SetError($__LO_STATUS_SUCCESS, 2, $sFilePath)
			If ($iPathSearch > 0) Then $sFilePath = "file:///" & $sFilePath

			For $i = 0 To (UBound($asURLReplace) - 1)
				$sFilePath = StringReplace($sFilePath, $asURLReplace[$i][0], $asURLReplace[$i][1])
				Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
			Next
			Return SetError($__LO_STATUS_SUCCESS, 2, $sFilePath)

		Case $LOB_PATHCONV_PCPATH_RETURN
			If ($iPathSearch > 0) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sFilePath)
			If ($iFileSearch > 0) Then $sFilePath = StringReplace($sFilePath, "file:///", Null)

			For $i = 0 To (UBound($asURLReplace) - 1)
				$sFilePath = StringReplace($sFilePath, $asURLReplace[$i][1], $asURLReplace[$i][0])
				Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
			Next
			Return SetError($__LO_STATUS_SUCCESS, 1, $sFilePath)

	EndSwitch

EndFunc   ;==>_LOBase_PathConvert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_VersionGet
; Description ...: Retrieve the current Office version.
; Syntax ........: _LOBase_VersionGet([$bSimpleVersion = False[, $bReturnName = False]])
; Parameters ....: $bSimpleVersion      - [optional] a boolean value. Default is False. If True, returns a two digit version number, such as "7.3", else returns the complex version number, such as "7.3.2.4".
;                  $bReturnName         - [optional] a boolean value. Default is True. If True returns the Program Name, such as "LibreOffice", appended by the version, i.e. "LibreOffice 7.3".
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bSimpleVersion not a Boolean.
;                  @Error 1 @Extended 2 Return 0 = $bReturnName not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.configuration.ConfigurationProvider" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error setting property value.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returns the Office version in String format.
; Author ........: Laurent Godard as found in Andrew Pitonyak's book; Zizi64 as found on OpenOffice forum.
; Modified ......: donnyh13, modified for AutoIt compatibility and error checking.
; Remarks .......: From Macro code by Zizi64 found at: https://forum.openoffice.org/en/forum/viewtopic.php?t=91542&sid=7f452d65e58ac1cd3cc6063350b5ada0
;                  And Andrew Pitonyak in "Useful Macro Information For OpenOffice.org" Pages 49, 50.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_VersionGet($bSimpleVersion = False, $bReturnName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sAccess = "com.sun.star.configuration.ConfigurationAccess", $sVersionName, $sVersion, $sReturn
	Local $oSettings, $oConfigProvider
	Local $aParamArray[1]

	If Not IsBool($bSimpleVersion) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Local $oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oConfigProvider = $oServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
	If Not IsObj($oConfigProvider) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$aParamArray[0] = __LOBase_SetPropertyValue("nodepath", "/org.openoffice.Setup/Product")
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSettings = $oConfigProvider.createInstanceWithArguments($sAccess, $aParamArray)

	$sVersionName = $oSettings.getByName("ooName")

	$sVersion = ($bSimpleVersion) ? ($oSettings.getByName("ooSetupVersion")) : ($oSettings.getByName("ooSetupVersionAboutBox"))

	$sReturn = ($bReturnName) ? ($sVersionName & " " & $sVersion) : ($sVersion)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sReturn)
EndFunc   ;==>_LOBase_VersionGet
