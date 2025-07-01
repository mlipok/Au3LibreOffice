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
; _LOBase_ConvertColorFromLong
; _LOBase_ConvertColorToLong
; _LOBase_ConvertFromMicrometer
; _LOBase_ConvertToMicrometer
; _LOBase_DateStructCreate
; _LOBase_DateStructModify
; _LOBase_FontDescCreate
; _LOBase_FontDescEdit
; _LOBase_FontExists
; _LOBase_FontsGetNames
; _LOBase_FormatKeyCreate
; _LOBase_FormatKeyDelete
; _LOBase_FormatKeyExists
; _LOBase_FormatKeyGetStandard
; _LOBase_FormatKeyGetString
; _LOBase_FormatKeyList
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
; Name ..........: _LOBase_ConvertColorFromLong
; Description ...: Convert Long color code to Hex, RGB, HSB or CMYK.
; Syntax ........: _LOBase_ConvertColorFromLong([$iHex = Null[, $iRGB = Null[, $iHSB = Null[, $iCMYK = Null]]]])
; Parameters ....: $iHex                - [optional] an integer value. Default is Null. Convert Long Color Integer to Hexadecimal.
;                  $iRGB                - [optional] an integer value. Default is Null. Convert Long Color Integer to R.G.B.
;                  $iHSB                - [optional] an integer value. Default is Null. Convert Long Color Integer to H.S.B.
;                  $iCMYK               - [optional] an integer value. Default is Null. Convert Long Color Integer to C.M.Y.K.
; Return values .: Success: String or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = No parameters set.
;                  @Error 1 @Extended 2 Return 0 = No parameters set to an integer.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Long integer converted To Hexadecimal (as a String). (Without the "0x" prefix)
;                  @Error 0 @Extended 2 Return Array = Array containing Long integer converted To Red, Green, Blue,(RGB). $Array[0] = R, $Array[1] = G, etc.
;                  @Error 0 @Extended 3 Return Array = Array containing Long integer converted To Hue, Saturation, Brightness, (HSB). $Array[0] = H, $Array[1] = S, etc.
;                  @Error 0 @Extended 4 Return Array = Array containing Long integer converted To Cyan, Magenta, Yellow, Black, (CMYK). $Array[0] = C, $Array[1] = M, etc.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve a Hexadecimal color value, call the Long Color code in $iHex, To retrieve a R(ed)G(reen)B(lue) color value, call Null in $iHex, and call the Long color code into $iRGB, etc. for the other color types.
;                  Hex returns as a string variable, all others (RGB, HSB, CMYK) return an array.
;                  The Hexadecimal figure returned doesn't contain the usual "0x", as LibeOffice does not implement it in its numbering system.
; Related .......: _LOBase_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ConvertColorFromLong($iHex = Null, $iRGB = Null, $iHSB = Null, $iCMYK = Null)
	Local $nRed, $nGreen, $nBlue, $nResult, $nMaxRGB, $nMinRGB, $nHue, $nSaturation, $nBrightness, $nCyan, $nMagenta, $nYellow, $nBlack
	Local $dHex
	Local $aiReturn[0]

	If (@NumParams = 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	Select
		Case IsInt($iHex) ; Long TO Hex
			$nRed = BitAND(BitShift($iHex, 16), 0xff)
			$nGreen = BitAND(BitShift($iHex, 8), 0xff)
			$nBlue = BitAND($iHex, 0xff)

			$dHex = Hex($nRed, 2) & Hex($nGreen, 2) & Hex($nBlue, 2)

			Return SetError($__LO_STATUS_SUCCESS, 1, $dHex)

		Case IsInt($iRGB) ; Long to RGB
			$nRed = BitAND(BitShift($iRGB, 16), 0xff)
			$nGreen = BitAND(BitShift($iRGB, 8), 0xff)
			$nBlue = BitAND($iRGB, 0xff)

			ReDim $aiReturn[3]
			$aiReturn[0] = $nRed
			$aiReturn[1] = $nGreen
			$aiReturn[2] = $nBlue

			Return SetError($__LO_STATUS_SUCCESS, 2, $aiReturn)

		Case IsInt($iHSB) ; Long TO HSB

			$nRed = (Mod(($iHSB / 65536), 256)) / 255
			$nGreen = (Mod(($iHSB / 256), 256)) / 255
			$nBlue = (Mod($iHSB, 256)) / 255

			; get Max RGB Value
			$nResult = ($nRed > $nGreen) ? ($nRed) : ($nGreen)
			$nMaxRGB = ($nResult > $nBlue) ? ($nResult) : ($nBlue)
			; get Min RGB Value
			$nResult = ($nRed < $nGreen) ? ($nRed) : ($nGreen)
			$nMinRGB = ($nResult < $nBlue) ? ($nResult) : ($nBlue)

			; Determine Brightness
			$nBrightness = $nMaxRGB
			; Determine Hue
			$nHue = 0
			Select
				Case $nRed = $nGreen = $nBlue ; Red, Green, and Blue are equal.
					$nHue = 0

				Case ($nRed >= $nGreen) And ($nGreen >= $nBlue) ; Red Highest, Blue Lowest
					$nHue = (60 * (($nGreen - $nBlue) / ($nRed - $nBlue)))

				Case ($nRed >= $nBlue) And ($nBlue >= $nGreen) ; Red Highest, Green Lowest
					$nHue = (60 * (6 - (($nBlue - $nGreen) / ($nRed - $nGreen))))

				Case ($nGreen >= $nRed) And ($nRed >= $nBlue) ; Green Highest, Blue Lowest
					$nHue = (60 * (2 - (($nRed - $nBlue) / ($nGreen - $nBlue))))

				Case ($nGreen >= $nBlue) And ($nBlue >= $nRed) ; Green Highest, Red Lowest
					$nHue = (60 * (2 + (($nBlue - $nRed) / ($nGreen - $nRed))))

				Case ($nBlue >= $nGreen) And ($nGreen >= $nRed) ; Blue Highest, Red Lowest
					$nHue = (60 * (4 - (($nGreen - $nRed) / ($nBlue - $nRed))))

				Case ($nBlue >= $nRed) And ($nRed >= $nGreen) ; Blue Highest, Green Lowest
					$nHue = (60 * (4 + (($nRed - $nGreen) / ($nBlue - $nGreen))))
			EndSelect

			; Determine Saturation
			$nSaturation = ($nMaxRGB = 0) ? (0) : (($nMaxRGB - $nMinRGB) / $nMaxRGB)

			$nHue = ($nHue > 0) ? (Round($nHue)) : (0)
			$nSaturation = Round(($nSaturation * 100))
			$nBrightness = Round(($nBrightness * 100))

			ReDim $aiReturn[3]
			$aiReturn[0] = $nHue
			$aiReturn[1] = $nSaturation
			$aiReturn[2] = $nBrightness

			Return SetError($__LO_STATUS_SUCCESS, 3, $aiReturn)

		Case IsInt($iCMYK) ; Long to CMYK

			$nRed = (Mod(($iCMYK / 65536), 256))
			$nGreen = (Mod(($iCMYK / 256), 256))
			$nBlue = (Mod($iCMYK, 256))

			$nRed = Round(($nRed / 255), 3)
			$nGreen = Round(($nGreen / 255), 3)
			$nBlue = Round(($nBlue / 255), 3)

			; get Max RGB Value
			$nResult = ($nRed > $nGreen) ? ($nRed) : ($nGreen)
			$nMaxRGB = ($nResult > $nBlue) ? ($nResult) : ($nBlue)

			$nBlack = (1 - $nMaxRGB)
			$nCyan = ((1 - $nRed - $nBlack) / (1 - $nBlack))
			$nMagenta = ((1 - $nGreen - $nBlack) / (1 - $nBlack))
			$nYellow = ((1 - $nBlue - $nBlack) / (1 - $nBlack))

			$nCyan = Round(($nCyan * 100))
			$nMagenta = Round(($nMagenta * 100))
			$nYellow = Round(($nYellow * 100))
			$nBlack = Round(($nBlack * 100))

			ReDim $aiReturn[4]
			$aiReturn[0] = $nCyan
			$aiReturn[1] = $nMagenta
			$aiReturn[2] = $nYellow
			$aiReturn[3] = $nBlack

			Return SetError($__LO_STATUS_SUCCESS, 4, $aiReturn)

		Case Else

			Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; no parameters set to an integer
	EndSelect
EndFunc   ;==>_LOBase_ConvertColorFromLong

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ConvertColorToLong
; Description ...: Convert Hex, RGB, HSB or CMYK to Long color code.
; Syntax ........: _LOBase_ConvertColorToLong([$vVal1 = Null[, $vVal2 = Null[, $vVal3 = Null[, $vVal4 = Null]]]])
; Parameters ....: $vVal1               - [optional] a variant value. Default is Null. See remarks.
;                  $vVal2               - [optional] a variant value. Default is Null. See remarks.
;                  $vVal3               - [optional] a variant value. Default is Null. See remarks.
;                  $vVal4               - [optional] a variant value. Default is Null. See remarks.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = No parameters set.
;                  @Error 1 @Extended 2 Return 0 = One parameter called, but not in String format(Hex).
;                  @Error 1 @Extended 3 Return 0 = Hex parameter contains non Hex characters.
;                  @Error 1 @Extended 4 Return 0 = Hex parameter not 6 characters long.
;                  @Error 1 @Extended 5 Return 0 = Hue parameter contains more than just digits.
;                  @Error 1 @Extended 6 Return 0 = Saturation parameter contains more than just digits.
;                  @Error 1 @Extended 7 Return 0 = Brightness parameter contains more than just digits.
;                  @Error 1 @Extended 8 Return 0 = Three parameters called but not all Integers (RGB) and not all Strings (HSB).
;                  @Error 1 @Extended 9 Return 0 = Four parameters called but not all Integers(CMYK).
;                  @Error 1 @Extended 10 Return 0 = Too many or too few parameters called.
;                  --Success--
;                  @Error 0 @Extended 1 Return Integer = Long Int. Color code converted from Hexadecimal.
;                  @Error 0 @Extended 2 Return Integer = Long Int. Color code converted from Red, Green, Blue, (RGB).
;                  @Error 0 @Extended 3 Return Integer = Long Int. Color code converted from (H)ue, (S)aturation, (B)rightness,
;                  @Error 0 @Extended 4 Return Integer = Long Int. Color code converted from (C)yan, (M)agenta, (Y)ellow, Blac(k)
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To Convert a Hex(adecimal) color code, call the Hex code in $vVal1 in String Format.
;                  To convert a R(ed) G(reen) B(lue color, call R value in $vVal1 as an Integer, G in $vVal2 as an Integer, and B in $vVal3 as an Integer.
;                  To convert a H(ue) S(aturation) B(rightness) color, call H in $vVal1 as a String, S in $vVal2 as a String, and B in $vVal3 as a string.
;                  To convert C(yan) M(agenta) Y(ellow) Blac(k) call C in $vVal1 as an Integer, M in $vVal2 as an Integer, Y in $vVal3 as an Integer, and K in $vVal4 as an Integer format.
;                  The Hexadecimal figure entered cannot contain the usual "0x", as LibeOffice does not implement it in its numbering system.
; Related .......: _LOBase_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ConvertColorToLong($vVal1 = Null, $vVal2 = Null, $vVal3 = Null, $vVal4 = Null) ; RGB = Int, CMYK = Int, HSB = String, Hex = String.
	Local Const $STR_STRIPALL = 8
	Local $iRed, $iGreen, $iBlue, $iLong, $iHue, $iSaturation, $iBrightness
	Local $dHex
	Local $nMaxRGB, $nMinRGB, $nChroma, $nHuePre, $nCyan, $nMagenta, $nYellow, $nBlack

	If (@NumParams = 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	Switch @NumParams
		Case 1 ; Hex
			If Not IsString($vVal1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; not a string
			$vVal1 = StringStripWS($vVal1, $STR_STRIPALL)
			$dHex = $vVal1

			; From Hex to RGB
			If (StringLen($dHex) = 6) Then
				If StringRegExp($dHex, "[^0-9a-fA-F]") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; $dHex contains non Hex characters.

				$iRed = BitAND(BitShift("0x" & $dHex, 16), 0xFF)
				$iGreen = BitAND(BitShift("0x" & $dHex, 8), 0xFF)
				$iBlue = BitAND("0x" & $dHex, 0xFF)

				$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

				Return SetError($__LO_STATUS_SUCCESS, 1, $iLong) ; Long from Hex

			Else

				Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Wrong length of string.
			EndIf

		Case 3 ; RGB and HSB; HSB is all strings, RGB all Integers.
			If (IsInt($vVal1) And IsInt($vVal2) And IsInt($vVal3)) Then ; RGB
				$iRed = $vVal1
				$iGreen = $vVal2
				$iBlue = $vVal3

				; RGB to Long
				$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

				Return SetError($__LO_STATUS_SUCCESS, 2, $iLong) ; Long from RGB

			ElseIf IsString($vVal1) And IsString($vVal2) And IsString($vVal3) Then ; Hue Saturation and Brightness (HSB)

				; HSB to RGB
				$vVal1 = StringStripWS($vVal1, $STR_STRIPALL)
				$vVal2 = StringStripWS($vVal2, $STR_STRIPALL)
				$vVal3 = StringStripWS($vVal3, $STR_STRIPALL) ; Strip WS so I can check string length in HSB conversion.

				$iHue = Number($vVal1)
				If (StringLen($vVal1)) <> (StringLen($iHue)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; String contained more than just digits
				$iSaturation = Number($vVal2)
				If (StringLen($vVal2)) <> (StringLen($iSaturation)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; String contained more than just digits
				$iBrightness = Number($vVal3)
				If (StringLen($vVal3)) <> (StringLen($iBrightness)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0) ; String contained more than just digits

				$nMaxRGB = ($iBrightness / 100)
				$nChroma = (($iSaturation / 100) * ($iBrightness / 100))
				$nMinRGB = ($nMaxRGB - $nChroma)
				$nHuePre = ($iHue >= 300) ? (($iHue - 360) / 60) : ($iHue / 60)

				Switch $nHuePre
					Case (-1) To 1.0
						$iRed = $nMaxRGB
						If $nHuePre < 0 Then
							$iGreen = $nMinRGB
							$iBlue = ($iGreen - $nHuePre * $nChroma)

						Else
							$iBlue = $nMinRGB
							$iGreen = ($iBlue + $nHuePre * $nChroma)
						EndIf

					Case 1.1 To 3.0
						$iGreen = $nMaxRGB
						If (($nHuePre - 2) < 0) Then
							$iBlue = $nMinRGB
							$iRed = ($iBlue - ($nHuePre - 2) * $nChroma)

						Else
							$iRed = $nMinRGB
							$iBlue = ($iRed + ($nHuePre - 2) * $nChroma)
						EndIf

					Case 3.1 To 5
						$iBlue = $nMaxRGB
						If (($nHuePre - 4) < 0) Then
							$iRed = $nMinRGB
							$iGreen = ($iRed - ($nHuePre - 4) * $nChroma)

						Else
							$iGreen = $nMinRGB
							$iRed = ($iGreen + ($nHuePre - 4) * $nChroma)
						EndIf
				EndSwitch

				$iRed = Round(($iRed * 255))
				$iGreen = Round(($iGreen * 255))
				$iBlue = Round(($iBlue * 255))

				$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

				Return SetError($__LO_STATUS_SUCCESS, 3, $iLong) ; Return Long from HSB

			Else

				Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0) ; Wrong parameters
			EndIf

		Case 4 ; CMYK
			If Not (IsInt($vVal1) And IsInt($vVal2) And IsInt($vVal3) And IsInt($vVal4)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0) ; CMYK not integers.

			; CMYK to RGB
			$nCyan = ($vVal1 / 100)
			$nMagenta = ($vVal2 / 100)
			$nYellow = ($vVal3 / 100)
			$nBlack = ($vVal4 / 100)

			$iRed = Round((255 * (1 - $nBlack) * (1 - $nCyan)))
			$iGreen = Round((255 * (1 - $nBlack) * (1 - $nMagenta)))
			$iBlue = Round((255 * (1 - $nBlack) * (1 - $nYellow)))

			$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

			Return SetError($__LO_STATUS_SUCCESS, 4, $iLong) ; Long from CMYK

		Case Else

			Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0) ; wrong number of Parameters
	EndSwitch
EndFunc   ;==>_LOBase_ConvertColorToLong

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ConvertFromMicrometer
; Description ...: Convert from Micrometer to Inch, Centimeter, Millimeter, or Printer's Points.
; Syntax ........: _LOBase_ConvertFromMicrometer([$nInchOut = Null[, $nCentimeterOut = Null[, $nMillimeterOut = Null[, $nPointsOut = Null]]]])
; Parameters ....: $nInchOut            - [optional] a general number value. Default is Null. The Micrometers to convert to Inches. See remarks.
;                  $nCentimeterOut      - [optional] a general number value. Default is Null. The Micrometers to convert to Centimeters. See remarks.
;                  $nMillimeterOut      - [optional] a general number value. Default is Null. The Micrometers to convert to Millimeters. See remarks.
;                  $nPointsOut          - [optional] a general number value. Default is Null. The Micrometers to convert to Printer's Points. See remarks.
; Return values .: Success: Number
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $nInchOut not a number.
;                  @Error 1 @Extended 2 Return 0 = $nCentimeterOut not a number.
;                  @Error 1 @Extended 3 Return 0 = $nMillimeterOut not a number.
;                  @Error 1 @Extended 4 Return 0 = $nPointsOut not a number.
;                  @Error 1 @Extended 5 Return 0 = No parameters set to other than Null.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error converting from Micrometers to Inch.
;                  @Error 3 @Extended 2 Return 0 = Error converting from Micrometers to Centimeter.
;                  @Error 3 @Extended 3 Return 0 = Error converting from Micrometers to Millimeter.
;                  @Error 3 @Extended 4 Return 0 = Error converting from Micrometers to Printer's Points.
;                  --Success--
;                  @Error 0 @Extended 1 Return Number = Converted from Micrometers to Inch.
;                  @Error 0 @Extended 2 Return Number = Converted from Micrometers to Centimeter.
;                  @Error 0 @Extended 3 Return Number = Converted from Micrometers to Millimeter.
;                  @Error 0 @Extended 4 Return Number = Converted from Micrometers to Printer's Points.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To skip a parameter, set it to Null.
;                  If you are converting to Inches, place the Micrometers in $nInchOut, if converting to Millimeters, $nInchOut and $nCentimeter are set to Null, and $nMillimetersOut is set.
;                  A Micrometer is 1000th of a centimeter, and is used in almost all Libre Office functions that contain a measurement parameter.
; Related .......: _LOBase_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ConvertFromMicrometer($nInchOut = Null, $nCentimeterOut = Null, $nMillimeterOut = Null, $nPointsOut = Null)
	Local $nReturnValue

	If ($nInchOut <> Null) Then
		If Not IsNumber($nInchOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
		$nReturnValue = __LOBase_UnitConvert($nInchOut, $__LOCONST_CONVERT_UM_INCH)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $nReturnValue)
	EndIf

	If ($nCentimeterOut <> Null) Then
		If Not IsNumber($nCentimeterOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$nReturnValue = __LOBase_UnitConvert($nCentimeterOut, $__LOCONST_CONVERT_UM_CM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 2, $nReturnValue)
	EndIf

	If ($nMillimeterOut <> Null) Then
		If Not IsNumber($nMillimeterOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$nReturnValue = __LOBase_UnitConvert($nMillimeterOut, $__LOCONST_CONVERT_UM_MM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Return SetError($__LO_STATUS_SUCCESS, 3, $nReturnValue)
	EndIf

	If ($nPointsOut <> Null) Then
		If Not IsNumber($nPointsOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$nReturnValue = __LOBase_UnitConvert($nPointsOut, $__LOCONST_CONVERT_UM_PT)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Return SetError($__LO_STATUS_SUCCESS, 4, $nReturnValue)
	EndIf

	Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; NO Unit set.
EndFunc   ;==>_LOBase_ConvertFromMicrometer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ConvertToMicrometer
; Description ...: Convert from Inch, Centimeter, Millimeter, or Printer's Points to Micrometer.
; Syntax ........: _LOBase_ConvertToMicrometer([$nInchIn = Null[, $nCentimeterIn = Null[, $nMillimeterIn = Null[, $nPointsIn = Null]]]])
; Parameters ....: $nInchIn             - [optional] a general number value. Default is Null. The Inches to convert to Micrometers. See remarks.
;                  $nCentimeterIn       - [optional] a general number value. Default is Null. The Centimeters to convert to Micrometers. See remarks.
;                  $nMillimeterIn       - [optional] a general number value. Default is Null. The Millimeters to convert to Micrometers. See remarks.
;                  $nPointsIn           - [optional] a general number value. Default is Null. The Printer's Points to convert to Micrometers. See remarks.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $nInchIn not a number.
;                  @Error 1 @Extended 2 Return 0 = $nCentimeterIn not a number.
;                  @Error 1 @Extended 3 Return 0 = $nMillimeterIn not a number.
;                  @Error 1 @Extended 4 Return 0 = $nPointsIn not a number.
;                  @Error 1 @Extended 5 Return 0 = No parameters set to other than Null.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error converting from Inches to Micrometers.
;                  @Error 3 @Extended 2 Return 0 = Error converting from Centimeters to Micrometers.
;                  @Error 3 @Extended 3 Return 0 = Error converting from Millimeters to Micrometers.
;                  @Error 3 @Extended 4 Return 0 = Error converting from Printer's Points to Micrometers.
;                  --Success--
;                  @Error 0 @Extended 1 Return Integer = Converted Inches to Micrometers.
;                  @Error 0 @Extended 2 Return Integer = Converted Centimeters to Micrometers.
;                  @Error 0 @Extended 3 Return Integer = Converted Millimeters to Micrometers.
;                  @Error 0 @Extended 4 Return Integer = Converted Printer's Points to Micrometers.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To skip a parameter, set it to Null. If you are converting from Inches, call inches in $nInchIn, if converting from Centimeters, $nInchIn is called with Null, and $nCentimeters is set.
;                  A Micrometer is 1000th of a centimeter, and is used in almost all Libre Office functions that contain a measurement parameter.
; Related .......: _LOBase_ConvertFromMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ConvertToMicrometer($nInchIn = Null, $nCentimeterIn = Null, $nMillimeterIn = Null, $nPointsIn = Null)
	Local $nReturnValue

	If ($nInchIn <> Null) Then
		If Not IsNumber($nInchIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
		$nReturnValue = __LOBase_UnitConvert($nInchIn, $__LOCONST_CONVERT_INCH_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $nReturnValue)
	EndIf

	If ($nCentimeterIn <> Null) Then
		If Not IsNumber($nCentimeterIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$nReturnValue = __LOBase_UnitConvert($nCentimeterIn, $__LOCONST_CONVERT_CM_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 2, $nReturnValue)
	EndIf

	If ($nMillimeterIn <> Null) Then
		If Not IsNumber($nMillimeterIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$nReturnValue = __LOBase_UnitConvert($nMillimeterIn, $__LOCONST_CONVERT_MM_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Return SetError($__LO_STATUS_SUCCESS, 3, $nReturnValue)
	EndIf

	If ($nPointsIn <> Null) Then
		If Not IsNumber($nPointsIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$nReturnValue = __LOBase_UnitConvert($nPointsIn, $__LOCONST_CONVERT_PT_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Return SetError($__LO_STATUS_SUCCESS, 4, $nReturnValue)
	EndIf

	Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; NO Unit set.
EndFunc   ;==>_LOBase_ConvertToMicrometer

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
; Name ..........: _LOBase_FontDescCreate
; Description ...: Create a Font Descriptor Map.
; Syntax ........: _LOBase_FontDescCreate([$sFontName = ""[, $iWeight = $LOB_WEIGHT_DONT_KNOW[, $iSlant = $LOB_POSTURE_DONTKNOW[, $nSize = 0[, $iColor = $LOB_COLOR_OFF[, $iUnderlineStyle = $LOB_UNDERLINE_DONT_KNOW[, $iUnderlineColor = $LOB_COLOR_OFF[, $iStrikelineStyle = $LOB_STRIKEOUT_DONT_KNOW[, $bIndividualWords = False[, $iRelief = $LOB_RELIEF_NONE[, $iCase = $LOB_CASEMAP_NONE[, $bHidden = False[, $bOutline = False[, $bShadow = False]]]]]]]]]]]]]])
; Parameters ....: $sFontName           - [optional] a string value. Default is "". The Font name.
;                  $iWeight             - [optional] an integer value (0-200). Default is $LOB_WEIGHT_DONT_KNOW. The Font weight. See Constants $LOB_WEIGHT_* as defined in LibreOfficeBase_Constants.au3.
;                  $iSlant              - [optional] an integer value (0-5). Default is $LOB_POSTURE_DONTKNOW. The Font italic setting. See Constants $LOB_POSTURE_* as defined in LibreOfficeBase_Constants.au3.
;                  $nSize               - [optional] a general number value. Default is 0. The Font size.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is $LOB_COLOR_OFF. The Font Color in Long Integer format, can be a custom value, or one of the constants, $LOB_COLOR_* as defined in LibreOfficeBase_Constants.au3. Set to $LOB_COLOR_OFF(-1) for Auto color.
;                  $iUnderlineStyle     - [optional] an integer value (0-18). Default is $LOB_UNDERLINE_DONT_KNOW. The Font underline Style. See Constants $LOB_UNDERLINE_* as defined in LibreOfficeBase_Constants.au3.
;                  $iUnderlineColor     - [optional] an integer value (-1-16777215). Default is $LOB_COLOR_OFF. The Font Underline color in Long Integer format, can be a custom value, or one of the constants, $LOB_COLOR_* as defined in LibreOfficeBase_Constants.au3. Set to $LOB_COLOR_OFF(-1) for Auto color.
;                  $iStrikelineStyle    - [optional] an integer value (0-6). Default is $LOB_STRIKEOUT_DONT_KNOW. The Strikeout line style. See Constants $LOB_STRIKEOUT_* as defined in LibreOfficeBase_Constants.au3.
;                  $bIndividualWords    - [optional] a boolean value. Default is False. If True, only individual words are underlined.
;                  $iRelief             - [optional] an integer value (0-2). Default is $LOB_RELIEF_NONE. The Font relief style. See Constants $LOB_RELIEF_* as defined in LibreOfficeBase_Constants.au3.
;                  $iCase               - [optional] an integer value (0-4). Default is $LOB_CASEMAP_NONE. The Character Case Style. See Constants, $LOB_CASEMAP_* as defined in LibreOfficeBase_Constants.au3.
;                  $bHidden             - [optional] a boolean value. Default is False. If True, the Characters are hidden.
;                  $bOutline            - [optional] a boolean value. Default is False. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] a boolean value. Default is False. If True, the characters have a shadow.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 2 Return 0 = Font called in $sFontName not found.
;                  @Error 1 @Extended 3 Return 0 = $iWeight not an Integer, less than 0 or greater than 200. See Constants $LOB_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iSlant not an Integer, less than 0 or greater than 5. See Constants $LOB_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $nSize not a number.
;                  @Error 1 @Extended 6 Return 0 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 7 Return 0 = $iUnderlineStyle not an Integer, less than 0 or greater than 18. See Constants $LOB_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iUnderlineColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 9 Return 0 = $iStrikelineStyle not an Integer, less than 0 or greater than 6. See Constants $LOB_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $bIndividualWords not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $iRelief not an Integer, less than 0 or greater than 2. See Constants $LOB_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 12 Return 0 = $iCase not an Integer, less than 0 or greater than 4. See Constants, $LOB_CASEMAP_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $bOutline not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $bShadow not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Returning the created Map Font Descriptor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FontDescCreate($sFontName = "", $iWeight = $LOB_WEIGHT_DONT_KNOW, $iSlant = $LOB_POSTURE_DONTKNOW, $nSize = 0, $iColor = $LOB_COLOR_OFF, $iUnderlineStyle = $LOB_UNDERLINE_DONT_KNOW, $iUnderlineColor = $LOB_COLOR_OFF, $iStrikelineStyle = $LOB_STRIKEOUT_DONT_KNOW, $bIndividualWords = False, $iRelief = $LOB_RELIEF_NONE, $iCase = $LOB_CASEMAP_NONE, $bHidden = False, $bOutline = False, $bShadow = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $mFontDesc[]

	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not _LOBase_FontExists($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOBase_IntIsBetween($iWeight, $LOB_WEIGHT_DONT_KNOW, $LOB_WEIGHT_BLACK) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LOBase_IntIsBetween($iSlant, $LOB_POSTURE_NONE, $LOB_POSTURE_REV_ITALIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsNumber($nSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not __LOBase_IntIsBetween($iColor, $LOB_COLOR_OFF, $LOB_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not __LOBase_IntIsBetween($iUnderlineStyle, $LOB_UNDERLINE_NONE, $LOB_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not __LOBase_IntIsBetween($iUnderlineColor, $LOB_COLOR_OFF, $LOB_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If Not __LOBase_IntIsBetween($iStrikelineStyle, $LOB_STRIKEOUT_NONE, $LOB_STRIKEOUT_X) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If Not IsBool($bIndividualWords) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
	If Not __LOBase_IntIsBetween($iRelief, $LOB_RELIEF_NONE, $LOB_RELIEF_ENGRAVED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
	If Not __LOBase_IntIsBetween($iCase, $LOB_CASEMAP_NONE, $LOB_CASEMAP_SM_CAPS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)
	If Not IsBool($bOutline) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)
	If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

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
	$mFontDesc.CharCaseMap = $iCase
	$mFontDesc.CharHidden = $bHidden
	$mFontDesc.CharContoured = $bOutline
	$mFontDesc.CharShadowed = $bShadow

	Return SetError($__LO_STATUS_SUCCESS, 0, $mFontDesc)
EndFunc   ;==>_LOBase_FontDescCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FontDescEdit
; Description ...: Set or Retrieve Font Descriptor settings.
; Syntax ........: _LOBase_FontDescEdit(ByRef $mFontDesc[, $sFontName = Null[, $iWeight = Null[, $iSlant = Null[, $nSize = Null[, $iColor = Null[, $iUnderlineStyle = Null[, $iUnderlineColor = Null[, $iStrikelineStyle = Null[, $bIndividualWords = Null[, $iRelief = Null[, $iCase = Null[, $bHidden = Null[, $bOutline = Null[, $bShadow = Null]]]]]]]]]]]]]])
; Parameters ....: $mFontDesc           - [in/out] a map. A Font descriptor Map as returned from a _LOBase_FontDescCreate, or control property return function.
;                  $sFontName           - [optional] a string value. Default is Null. The Font name.
;                  $iWeight             - [optional] an integer value (0-200). Default is Null. The Font weight. See Constants $LOB_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iSlant              - [optional] an integer value (0-5). Default is Null. The Font italic setting. See Constants $LOB_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $nSize               - [optional] a general number value. Default is Null. The Font size.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is Null. The Font Color in Long Integer format, can be a custom value, or one of the constants, $LOB_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOB_COLOR_OFF(-1) for Auto color.
;                  $iUnderlineStyle     - [optional] an integer value (0-18). Default is Null. The Font underline Style. See Constants $LOB_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iUnderlineColor     - [optional] an integer value (-1-16777215). Default is Null.
;                  $iStrikelineStyle    - [optional] an integer value (0-6). Default is Null. The Strikeout line style. See Constants $LOB_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bIndividualWords    - [optional] a boolean value. Default is Null. If True, only individual words are underlined.
;                  $iRelief             - [optional] an integer value (0-2). Default is Null. The Font relief style. See Constants $LOB_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iCase               - [optional] an integer value (0-4). Default is Null. The Character Case Style. See Constants, $LOB_CASEMAP_* as defined in LibreOfficeBase_Constants.au3.
;                  $bHidden             - [optional] a boolean value. Default is Null. If True, the Characters are hidden.
;                  $bOutline            - [optional] a boolean value. Default is False. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] a boolean value. Default is False. If True, the characters have a shadow.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $mFontDesc not a Map.
;                  @Error 1 @Extended 2 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 3 Return 0 = Font called in $sFontName not found.
;                  @Error 1 @Extended 4 Return 0 = $iWeight not an Integer, less than 0 or greater than 200. See Constants $LOB_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iSlant not an Integer, less than 0 or greater than 5. See Constants $LOB_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $nSize not a number.
;                  @Error 1 @Extended 7 Return 0 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 8 Return 0 = $iUnderlineStyle not an Integer, less than 0 or greater than 18. See Constants $LOB_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $iUnderlineColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 10 Return 0 = $iStrikelineStyle not an Integer, less than 0 or greater than 6. See Constants $LOB_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = $bIndividualWords not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iRelief not an Integer, less than 0 or greater than 2. See Constants $LOB_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $iCase not an Integer, less than 0 or greater than 4. See Constants, $LOB_CASEMAP_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 14 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $bOutline not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $bShadow not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 14 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FontDescEdit(ByRef $mFontDesc, $sFontName = Null, $iWeight = Null, $iSlant = Null, $nSize = Null, $iColor = Null, $iUnderlineStyle = Null, $iUnderlineColor = Null, $iStrikelineStyle = Null, $bIndividualWords = Null, $iRelief = Null, $iCase = Null, $bHidden = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avFont[14]

	If Not IsMap($mFontDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOBase_VarsAreNull($sFontName, $iWeight, $iSlant, $nSize, $iColor, $iUnderlineStyle, $iUnderlineColor, $iStrikelineStyle, $bIndividualWords, $iRelief, $iCase, $bHidden, $bOutline, $bShadow) Then
		__LOBase_ArrayFill($avFont, $mFontDesc.CharFontName, $mFontDesc.CharWeight, $mFontDesc.CharPosture, $mFontDesc.CharHeight, $mFontDesc.CharColor, $mFontDesc.CharUnderline, _
				$mFontDesc.CharUnderlineColor, $mFontDesc.CharStrikeout, $mFontDesc.CharWordMode, $mFontDesc.CharRelief, $mFontDesc.CharCaseMap, $mFontDesc.CharHidden, _
				$mFontDesc.CharContoured, $mFontDesc.CharShadowed)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFont)
	EndIf

	If ($sFontName <> Null) Then
		If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		If Not _LOBase_FontExists($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$mFontDesc.CharFontName = $sFontName
	EndIf

	If ($iWeight <> Null) Then
		If Not __LOBase_IntIsBetween($iWeight, $LOB_WEIGHT_DONT_KNOW, $LOB_WEIGHT_BLACK) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$mFontDesc.CharWeight = $iWeight
	EndIf

	If ($iSlant <> Null) Then
		If Not __LOBase_IntIsBetween($iSlant, $LOB_POSTURE_NONE, $LOB_POSTURE_REV_ITALIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$mFontDesc.CharPosture = $iSlant
	EndIf

	If ($nSize <> Null) Then
		If Not IsNumber($nSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$mFontDesc.CharHeight = $nSize
	EndIf

	If ($iColor <> Null) Then
		If Not __LOBase_IntIsBetween($iColor, $LOB_COLOR_OFF, $LOB_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$mFontDesc.CharColor = $iColor
	EndIf

	If ($iUnderlineStyle <> Null) Then
		If Not __LOBase_IntIsBetween($iUnderlineStyle, $LOB_UNDERLINE_NONE, $LOB_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$mFontDesc.CharUnderline = $iUnderlineStyle
	EndIf

	If ($iUnderlineColor <> Null) Then
		If Not __LOBase_IntIsBetween($iUnderlineColor, $LOB_COLOR_OFF, $LOB_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		$mFontDesc.CharUnderlineColor = $iUnderlineColor
	EndIf

	If ($iStrikelineStyle <> Null) Then
		If Not __LOBase_IntIsBetween($iStrikelineStyle, $LOB_STRIKEOUT_NONE, $LOB_STRIKEOUT_X) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		$mFontDesc.CharStrikeout = $iStrikelineStyle
	EndIf

	If ($bIndividualWords <> Null) Then
		If Not IsBool($bIndividualWords) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
		$mFontDesc.CharWordMode = $bIndividualWords
	EndIf

	If ($iRelief <> Null) Then
		If Not __LOBase_IntIsBetween($iRelief, $LOB_RELIEF_NONE, $LOB_RELIEF_ENGRAVED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)
		$mFontDesc.CharRelief = $iRelief
	EndIf

	If ($iCase <> Null) Then
		If Not __LOBase_IntIsBetween($iCase, $LOB_CASEMAP_NONE, $LOB_CASEMAP_SM_CAPS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)
		$mFontDesc.CharCaseMap = $iCase
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)
		$mFontDesc.CharHidden = $bHidden
	EndIf

	If ($bOutline <> Null) Then
		If Not IsBool($bOutline) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)
		$mFontDesc.CharContoured = $bOutline
	EndIf

	If ($bShadow <> Null) Then
		If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)
		$mFontDesc.CharShadowed = $bShadow
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FontDescEdit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FontExists
; Description ...: Tests whether a Document has a specific font available by name.
; Syntax ........: _LOBase_FontExists($sFontName)
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
Func _LOBase_FontExists($sFontName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $oServiceManager, $oDesktop, $oDoc
	Local $atProperties[1]

	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$atProperties[0] = __LOBase_SetPropertyValue("Hidden", True)
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

		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOBase_FontExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FontsGetNames
; Description ...: Retrieve an array of currently available fonts.
; Syntax ........: _LOBase_FontsGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen, _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
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
;                  The third column (Array[1][2]) contains the Font weight (Bold) See Constants, $LOB_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3;
;                  The fourth column (Array[1][3]) contains the font slant (Italic) See constants, $LOB_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FontsGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
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
		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($atFonts), $asFonts)
EndFunc   ;==>_LOBase_FontsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyCreate
; Description ...: Create a Format Key.
; Syntax ........: _LOBase_FormatKeyCreate(ByRef $oReportDoc, $sFormat)
; Parameters ....: $oReportDoc          - [in/out] an object. A Document object returned by a previous _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $sFormat             - a string value. The format key String to create.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFormat not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Attempted to Create or Retrieve the Format key, but failed.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Format Key was successfully created, returning Format Key integer.
;                  @Error 0 @Extended 1 Return Integer = Success. Format Key already existed, returning Format Key integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_FormatKeyDelete, _LOBase_FormatKeyGetStandard
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyCreate(ByRef $oReportDoc, $sFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFormatKey
	Local $tLocale
	Local $oFormats

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$tLocale = __LOBase_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oReportDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	$iFormatKey = $oFormats.queryKey($sFormat, $tLocale, False)
	If ($iFormatKey > -1) Then Return SetError($__LO_STATUS_SUCCESS, 1, $iFormatKey) ; Format already existed
	$iFormatKey = $oFormats.addNew($sFormat, $tLocale)
	If ($iFormatKey > -1) Then Return SetError($__LO_STATUS_SUCCESS, 0, $iFormatKey) ; Format created

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Failed to create or retrieve Format
EndFunc   ;==>_LOBase_FormatKeyCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyDelete
; Description ...: Delete a User-Created Format Key from a Document.
; Syntax ........: _LOBase_FormatKeyDelete(ByRef $oReportDoc, $iFormatKey)
; Parameters ....: $oReportDoc          - [in/out] an object. A Document object returned by a previous _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $iFormatKey          - an integer value. The User-Created format Key to delete.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 4 Return 0 = Format Key called in $iFormatKey not found in Document.
;                  @Error 1 @Extended 5 Return 0 = Format Key called in $iFormatKey not User-Created.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to delete key.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Format Key was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_FormatKeyList, _LOBase_FormatKeyCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyDelete(ByRef $oReportDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not _LOBase_FormatKeyExists($oReportDoc, $iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Key not found.

	$oFormats = $oReportDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($oFormats.getbykey($iFormatKey).UserDefined() = False) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; Key not User Created.

	$oFormats.removeByKey($iFormatKey)

	Return (_LOBase_FormatKeyExists($oReportDoc, $iFormatKey) = False) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0))
EndFunc   ;==>_LOBase_FormatKeyDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyExists
; Description ...: Check if a Document contains a certain Format Key.
; Syntax ........: _LOBase_FormatKeyExists(ByRef $oReportDoc, $iFormatKey[, $iFormatType = $LOB_FORMAT_KEYS_ALL])
; Parameters ....: $oReportDoc          - [in/out] an object. A Document object returned by a previous _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $iFormatKey          - an integer value. The Format Key to look for.
;                  $iFormatType         - [optional] an integer value (0-15881). Default is $LOB_FORMAT_KEYS_ALL. The Format Key type to search in. Values can be BitOr'd together. See Constants, $LOB_FORMAT_KEYS_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iFormatType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to obtain Array of Date/Time Formats.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the Format Key exists in document, True is Returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyExists(ByRef $oReportDoc, $iFormatKey, $iFormatType = $LOB_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys[0]
	Local $tLocale

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iFormatType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	$tLocale = __LOBase_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oReportDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	$aiFormatKeys = $oFormats.queryKeys($iFormatType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aiFormatKeys) - 1
		If ($aiFormatKeys[$i] = $iFormatKey) Then Return SetError($__LO_STATUS_SUCCESS, 0, True) ; Doc does contain format Key
		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, False) ; Doc does not contain format Key
EndFunc   ;==>_LOBase_FormatKeyExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyGetStandard
; Description ...: Retrieve the Standard Format for a specific Format Key Type.
; Syntax ........: _LOBase_FormatKeyGetStandard(ByRef $oReportDoc, $iFormatKeyType)
; Parameters ....: $oReportDoc          - [in/out] an object. A Document object returned by a previous _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $iFormatKeyType      - an integer value (1-8196). The Format Key type to retrieve the standard Format for. See Constants $LOB_FORMAT_KEYS_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $iFormatKeyType not an Integer, less than 1 or greater than 8196. See Constants $LOB_FORMAT_KEYS_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.lang.Locale" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Standard Format for the requested Format Key Type.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning the Standard Format for the requested Format Key Type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyGetStandard(ByRef $oReportDoc, $iFormatKeyType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $tLocale
	Local $iStandard

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOBase_IntIsBetween($iFormatKeyType, $LOB_FORMAT_KEYS_DEFINED, $LOB_FORMAT_KEYS_DURATION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tLocale = __LOBase_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oReportDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iStandard = $oFormats.getStandardFormat($iFormatKeyType, $tLocale)
	If Not IsInt($iStandard) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iStandard)
EndFunc   ;==>_LOBase_FormatKeyGetStandard

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyGetString
; Description ...: Retrieve a Format Key String.
; Syntax ........: _LOBase_FormatKeyGetString(ByRef $oReportDoc, $iFormatKey)
; Parameters ....: $oReportDoc          - [in/out] an object. A Document object returned by a previous _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $iFormatKey          - an integer value. The Format Key to retrieve the string for.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iFormatKey not found in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve requested Format Key Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning Format Key's Format String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_FormatKeyList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyGetString(ByRef $oReportDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormatKey

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not _LOBase_FormatKeyExists($oReportDoc, $iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	$oFormatKey = $oReportDoc.getNumberFormats().getByKey($iFormatKey)
	If Not IsObj($oFormatKey) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Key not found.

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFormatKey.FormatString())
EndFunc   ;==>_LOBase_FormatKeyGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyList
; Description ...: Retrieve an Array of Date/Time Format Keys.
; Syntax ........: _LOBase_FormatKeyList(ByRef $oReportDoc[, $bIsUser = False[, $bUserOnly = False[, $iFormatKeyType = $LOB_FORMAT_KEYS_ALL]]])
; Parameters ....: $oReportDoc          - [in/out] an object. A Document object returned by a previous _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $bIsUser             - [optional] a boolean value. Default is False. If True, Adds a third column to the return Array with a boolean, whether each Key is user-created or not.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True, only user-created Format Keys are returned.
;                  $iFormatKeyType      - [optional] an integer value (0-15881). Default is $LOB_FORMAT_KEYS_ALL. The Format Key type to retrieve an array of. Values can be BitOr'd together. See Constants, $LOB_FORMAT_KEYS_* as defined in LibreOfficeWriter_Constants.au3..
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $bIsUser not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bUserOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iFormatKeyType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve NumberFormats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to obtain Array of Format Keys.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning a 2 or three column Array, depending on current $bIsUser setting. See remarks. @Extended is set to the number of Keys returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Column One (Array[0][0]) will contain the Format Key integer,
;                  Column two (Array[0][1]) will contain the Format Key String,
;                  If $bIsUser is set to True, Column Three (Array[0][2]) will contain a Boolean, True if the Format Key is User created, else false.
; Related .......: _LOBase_FormatKeyDelete, _LOBase_FormatKeyGetString, _LOBase_FormatKeyGetStandard
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyList(ByRef $oReportDoc, $bIsUser = False, $bUserOnly = False, $iFormatKeyType = $LOB_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys
	Local $avFormats[0][3]
	Local $tLocale
	Local $iColumns = 3, $iCount = 0

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bIsUser) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	$iColumns = ($bIsUser = True) ? ($iColumns) : (2)

	If Not IsInt($iFormatKeyType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$tLocale = __LOBase_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oReportDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	$aiFormatKeys = $oFormats.queryKeys($iFormatKeyType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

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
		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	If ($bUserOnly = True) Then ReDim $avFormats[$iCount][$iColumns]

	Return SetError($__LO_STATUS_SUCCESS, UBound($avFormats), $avFormats)
EndFunc   ;==>_LOBase_FormatKeyList

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
