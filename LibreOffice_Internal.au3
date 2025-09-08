#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;~ #Tidy_Parameters=/sf /reel

#include-once

#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Internal functions for interacting with Libre Office.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __LO_DeleteTempReg
; __LO_InternalComErrorHandler
; __LO_ServiceManager
; __LO_SetPortableServiceManager
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_DeleteTempReg
; Description ...: Delete Temporary Registry entries used for connecting to Portable LO.
; Syntax ........: __LO_DeleteTempReg([$asRegKeys = Null])
; Parameters ....: $asRegKeys           - [optional] an array of strings. Default is Null. An array of Registry keys to Delete.
; Return values .: Success: 1, 2
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $asRegKeys not an Array.
;                  --Processing Errors--
;                  @Error 3 @Extended ? Return 0 = Error Deleting Registry key. @Extended set to number of errors encountered.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully stored Registry keys to delete.
;                  @Error 0 @Extended 0 Return 2 = Success. Successfully deleted Registry keys.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_DeleteTempReg($asRegKeys = Null)
	Local Static $asStaticKeys[0]
	Local $iError = 0
	Local Const $sHKCU = (@OSArch = "X86") ? ("HKCU") : ("HKCU64")

	If ($asRegKeys <> Null) Then
		If Not IsArray($asRegKeys) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		$asStaticKeys = $asRegKeys

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndIf

	For $sKey In $asStaticKeys
		RegDelete($sHKCU & $sKey)
		$iError = (@error > 0) ? ($iError + 1) : ($iError)
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROCESSING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 2))
EndFunc   ;==>__LO_DeleteTempReg

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LO_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _LO_ComError_UserFunction(Default)
	Local $vUserFunction, $avUserParams[2] = ["CallArgArray", $oComError]

	If IsArray($avUserFunction) Then
		$vUserFunction = $avUserFunction[0]
		ReDim $avUserParams[UBound($avUserFunction) + 1]
		For $i = 1 To UBound($avUserFunction) - 1
			$avUserParams[$i + 1] = $avUserFunction[$i]
		Next

	Else
		$vUserFunction = $avUserFunction
	EndIf
	If IsFunc($vUserFunction) Then
		Switch $vUserFunction
			Case ConsoleWrite
				ConsoleWrite("!--COM Error-Begin--" & @CRLF & _
						"Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline & @CRLF & _
						"!--COM-Error-End--" & @CRLF)

			Case MsgBox
				MsgBox(0, "COM Error", "Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline)

			Case Else
				Call($vUserFunction, $avUserParams)
		EndSwitch
	EndIf
EndFunc   ;==>__LO_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_ServiceManager
; Description ...: Set or Retrieve a stored Service Manager Object for use in the UDF.
; Syntax ........: __LO_ServiceManager([$oServiceManager = Null[, $bPortable = Null]])
; Parameters ....: $oServiceManager     - [optional] an object. Default is Null. A ServiceManager Object. Typically this is used to store a Portable Service Manager Object.
;                  $bPortable           - [optional] a boolean value. Default is Null. If True, a Portable LibreOffice ServiceManager will be stored.
; Return values .: Success: 1 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bPortable not a Boolean.
;                  @Error 1 @Extended 2 Return 0 = $oServiceManager not an Object.
;                  @Error 1 @Extended 3 Return 0 =Object called in $oServiceManager not a ServiceManager Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a ServiceManager.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared stored ServiceManager.
;                  @Error 0 @Extended 0 Return Object = Success. Returning ServiceManager Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_ServiceManager($oServiceManager = Null, $bPortable = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Static $oStaticServiceManager
	Local Static $bIsPortable = False

	If ($bPortable <> Null) Then
		If Not IsBool($bPortable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		$bIsPortable = $bPortable
	EndIf

	If ($oServiceManager <> Null) Then
		If ($oServiceManager = Default) Then ; Clear the saved Service Manager. This could be used in the case of switching from portable to installed.
			$oStaticServiceManager = Null

		Else
			If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
			If Not $oServiceManager.supportsService("com.sun.star.lang.ServiceManager") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$oStaticServiceManager = $oServiceManager
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndIf

	If IsObj($oStaticServiceManager) Then ; Test if the Object is still valid.
		If Not IsBool($oStaticServiceManager.supportsService("com.sun.star.lang.ServiceManager")) Then $oStaticServiceManager = Null
	EndIf

	If Not IsObj($oStaticServiceManager) Then
		If $bIsPortable Then
			; Try to create the ServiceManager again for the portable version.
			__LO_SetPortableServiceManager()

		Else ; Create a ServiceManager, for the installed version.
			$oStaticServiceManager = ObjCreate("com.sun.star.ServiceManager")
		EndIf
	EndIf

	If Not IsObj($oStaticServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $oStaticServiceManager)
EndFunc   ;==>__LO_ServiceManager

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_SetPortableServiceManager
; Description ...: Create and store a Portable LibreOffice ServiceManager Object.
; Syntax ........: __LO_SetPortableServiceManager([$sPortableLO_Path = Null])
; Parameters ....: $sPortableLO_Path    - [optional] a string value. Default is Null. A path to the Portable LibreOffice soffice.exe file.
; Return values .: Success: 1, 2
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sPortableLO_Path not a String.
;                  @Error 1 @Extended 2 Return 0 = Path called in $sPortableLO_Path doesn't exist.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.bridge.UnoUrlResolver" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = No stored Portable LibreOffice path.
;                  @Error 3 @Extended 2 Return 0 = Stored Portable LibreOffice path no longer exists.
;                  @Error 3 @Extended 3 Return 0 = Failed to add temporary Registry keys.
;                  @Error 3 @Extended 4 Return 0 = Failed to start Portable LibreOffice in listening mode.
;                  @Error 3 @Extended 5 Return 0 = Portable LibreOffice failed to start in listening mode
;                  @Error 3 @Extended 6 Return 0 = Failed to connect to Portable LibreOffice.
;                  @Error 3 @Extended 7 Return 0 = Failed to retrieve ServiceManager from Portable LibreOffice.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Created and stored Portable LibreOffice ServiceManager.
;                  @Error 0 @Extended 0 Return 2 = Success. Cleared stored Portable LibreOffice path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the COM Error "Binary URP bridge already disposed" is encountered, any running soffice.exe/soffice.bin processes need to be closed via TaskManager.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_SetPortableServiceManager($sPortableLO_Path = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Static $sStaticPortablePath = ""
	Local Static $bTempReg = False
	Local $iError = 0, $iPort = 2002, $iSocket, $iErrorRet
	Local $hTimer
	Local $oReg_ServiceManager, $oURLResolver, $oCompContext, $oPortable_ServiceManager
	Local Const $__eReg_LocalServer32 = 2 ; Make sure to sync this.
	Local Enum $__eReg_KeyName, $__eReg_ValueName, $__eReg_Type, $__eReg_Value
	Local Const $sHKCU = (@OSArch = "X86") ? ("HKCU") : ("HKCU64")
	Local Const $sHKLM = (@OSArch = "X86") ? ("HKLM") : ("HLM64")
	Local $asRegKeysMain[2] = ["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}", _
			"\Software\Classes\com.sun.star.ServiceManager"]
	Local $asRegKeys[11][4] = [["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}", "", "REG_SZ", "LibreOffice Service Manager (Ver 1.0)"], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}", "AppID", "REG_SZ", "{82154420-0FBF-11d4-8313-005004526AB4}"], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}\LocalServer32", "", "REG_SZ", ""], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}\NotInsertable", "", "REG_SZ", ""], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}\ProgID", "", "REG_SZ", "com.sun.star.ServiceManager.1"], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}\Programmable", "", "REG_SZ", ""], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}\VersionIndependentProgID", "", "REG_SZ", "com.sun.star.ServiceManager"], _
			["\Software\Classes\com.sun.star.ServiceManager", "", "REG_SZ", "LibreOffice Service Manager"], _
			["\Software\Classes\com.sun.star.ServiceManager\CLSID", "", "REG_SZ", "{82154420-0FBF-11d4-8313-005004526AB4}"], _
			["\Software\Classes\com.sun.star.ServiceManager\CurVer", "", "REG_SZ", "com.sun.star.ServiceManager.1"], _
			["\Software\Classes\com.sun.star.ServiceManager\NotInsertable", "", "REG_SZ", ""]]

	If ($sPortableLO_Path <> Null) Then
		If Not IsString($sPortableLO_Path) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		If ($sPortableLO_Path <> "") Then
			If Not FileExists($sPortableLO_Path) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

			$sStaticPortablePath = $sPortableLO_Path

		Else
			$sStaticPortablePath = $sPortableLO_Path
			__LO_ServiceManager(Default, False) ; Clear any stored ServiceManager, and set Boolean for Portable to False.

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf
	EndIf

	If ($sStaticPortablePath = "") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not FileExists($sStaticPortablePath) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Check to make sure the stored path to the LO File is still good.

	$asRegKeys[$__eReg_LocalServer32][$__eReg_Value] = $sStaticPortablePath & " --nodefault --nologo"

	RegRead($sHKLM & $asRegKeysMain[1], "") ; Classes are shared, so I only need to check one place.
	If @error Then ; Seems like no LibreOffice is installed, add temp Registry.
		For $i = 0 To UBound($asRegKeys) - 1
			RegWrite($sHKCU & $asRegKeys[$i][$__eReg_KeyName], $asRegKeys[$i][$__eReg_ValueName], $asRegKeys[$i][$__eReg_Type], $asRegKeys[$i][$__eReg_Value])
			$iError = (@error > 0) ? ($iError + 1) : ($iError)
		Next

		__LO_DeleteTempReg($asRegKeysMain) ; Set array of main Temp keys to delete.
		If ($iError > 0) Then ; If there was an error writing the Reg Keys, delete any that were written and return.
			__LO_DeleteTempReg()

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		EndIf

		OnAutoItExitRegister("__LO_DeleteTempReg")
		$bTempReg = True
	EndIf

	$oReg_ServiceManager = ObjCreate("com.sun.star.ServiceManager") ; Create a ServiceManager from the one registered in Registry.
	If Not IsObj($oReg_ServiceManager) Then
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf

		Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	TCPStartup()

	; Find an unused port.
	While (TCPConnect("127.0.0.1", $iPort) > 0)
		Sleep(10)
		$iPort += 1
	WEnd

	; Start LibreOffice Portable in listening mode
	Run('"' & $sStaticPortablePath & '" --headless --norestore --nologo --accept="socket,host=127.0.0.1,port=' & $iPort & ',tcpNoDelay=1;urp;"', "", @SW_HIDE)
	If (@error <> 0) Then
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf
		TCPShutdown()

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$hTimer = TimerInit()

	; Wait until LO is initialized and listening.
	Do
		Sleep(10)
		$iSocket = TCPConnect("127.0.0.1", $iPort)
		$iErrorRet = @error
	Until (($iErrorRet = 0) And ($iSocket > 0)) Or (TimerDiff($hTimer) > 15000) ; Timeout in 15 seconds.
	TCPShutdown()

	If ($iErrorRet > 0) Then ; Error initializing LO.
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	; Create the URL Resolver.
	$oURLResolver = $oReg_ServiceManager.createInstance("com.sun.star.bridge.UnoUrlResolver")
	If Not IsObj($oURLResolver) Then
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf

		Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	EndIf

	$oCompContext = $oURLResolver.resolve("uno:socket,host=localhost,port=" & $iPort & ",tcpNoDelay=1;urp;StarOffice.ComponentContext")

	If Not IsObj($oCompContext) Then
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
	EndIf

	$oPortable_ServiceManager = $oCompContext.getServiceManager() ; Get ServiceManager of Portable LO.
	If Not IsObj($oPortable_ServiceManager) Then
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)
	EndIf

	__LO_ServiceManager($oPortable_ServiceManager, True) ; Set the stored ServiceManager.

	If $bTempReg Then ; Clean up Temp Registry.
		__LO_DeleteTempReg()
		If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
		$bTempReg = False
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LO_SetPortableServiceManager
