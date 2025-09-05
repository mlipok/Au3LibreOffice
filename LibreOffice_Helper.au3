#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;~ #Tidy_Parameters=/sf /reel

#include-once

#include "LibreOffice_Constants.au3"
#include "LibreOffice_Internal.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Helper functions for using this UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LO_ComError_UserFunction
; _LO_InitializePortable
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LO_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
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
;                  Every COM Error will be passed to that function. The user can then read the following properties. (As Found in the COM Reference section in Autoit HelpFile.) Using the first parameter in the UserFunction.
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
Func _LO_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
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
EndFunc   ;==>_LO_ComError_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_InitializePortable
; Description ...: Setup Portable LibreOffice (Or Open Office) for use in this UDF. See remarks.
; Syntax ........: _LO_InitializePortable($sOfficePortablePath)
; Parameters ....: $sOfficePortablePath - a string value. The Path to the Portable LibreOffice/OpenOffice folder. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sOfficePortablePath not a String.
;                  @Error 1 @Extended 2 Return 0 = Folder called in $sOfficePortablePath does not contain the App folder. Perhaps wrong directory?
;                  @Error 1 @Extended 3 Return 0 = soffice.exe not found in $sOfficePortablePath\App\libreoffice\program\ or $sOfficePortablePath\App\openoffice\program\.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to clear stored Portable LO/OO ServiceManager.
;                  @Error 3 @Extended 2 Return 0 = Failed to initialize portable LibreOffice ServiceManager.
;                  @Error 3 @Extended 3 Return 0 = Failed to initialize portable OpenOffice ServiceManager.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Portable LibreOffice/OpenOffice ServiceManager successfully created and stored.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The path called in $sOfficePortablePath should be to the Portable LibreOffice folder containing the shortcuts to each element, and also the "App", "Data" and "Other" folders. e.g. C:\LibreOfficePortablePrevious
;                  This UDF hasn't been thoroughly tested using portable LibreOffice, but rather the Installed version. Make sure to test all things yourself!
;                  So far, the following method is the only way I've been able to successfully initialize the portable LibreOffice version from AutoIt. How it works is as follows:
;                  1a. If LibreOffice/OpenOffice is already installed on the system, an instance of com.sun.star.ServiceManager is started. __OR__
;                  1b. If LibreOffice/OpenOffice is NOT already installed on the system, a handful of temporary registry entries are created in HKEY_CURRENT_USER, to allow an instance of com.sun.star.ServiceManager to be started.
;                  2. The Portable LibreOffice/OpenOffice (LO/OO) is started in --headless mode as a listening server, using the flag --accept.
;                  3. The ServiceManager created in step 1 is used to create a UNO URL resolver, which is used to obtain a ServiceManager Object from the listening Portable LO/OO.
;                  4. The retrieved Portable LO/OO ServiceManager is stored as a static variable for future use. The Portable LO/OO path is also stored as a static variable in case I need to re-create the ServiceManager.
;                  5a. The ServiceManager created from the registry is no longer used.
;                  5b. The ServiceManager created from the registry is no longer used, and the temporary Registry entries created in HKEY_CURRENT_USER are (hopefully) deleted.
;                  If the COM error "Binary URP bridge already disposed" is encountered, all instances of soffice.exe or soffice.bin must be closed with TaskManager.
;                  If running this with an installed version of LibreOffice present the flag SingleAppInstance may need to be set to false in the "LibreOfficePortablePrevious.ini" [or similar name], found at: C:\LibreOfficePortablePrevious\App\AppInfo\Launcher\LibreOfficePortablePrevious.ini.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LO_InitializePortable($sOfficePortablePath)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsString($sOfficePortablePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sOfficePortablePath = StringRegExpReplace($sOfficePortablePath, "^\s+|\s+$", "") ; Strip beginning and ending spaces.

	If ($sOfficePortablePath <> "") And Not FileExists($sOfficePortablePath & "\App") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Make sure we're starting from the right folder.

	If ($sOfficePortablePath = "") Then
		__LO_SetPortableServiceManager($sOfficePortablePath)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ElseIf FileExists($sOfficePortablePath & "\App\libreoffice\program\soffice.exe") Then ; Check Libre path.
		__LO_SetPortableServiceManager($sOfficePortablePath & "\App\libreoffice\program\soffice.exe")
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	ElseIf FileExists($sOfficePortablePath & "\App\openoffice\program\soffice.exe") Then ; Check OpenOffice path.
		__LO_SetPortableServiceManager($sOfficePortablePath & "\App\openoffice\program\soffice.exe")
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Else
		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LO_InitializePortable
