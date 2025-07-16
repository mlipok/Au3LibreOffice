#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;~ #Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Constants.au3"
#include "LibreOfficeCalc_Internal.au3"

; Other includes for Calc

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Functions used for creating, modifying and retrieving data for use in various functions in LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_ComError_UserFunction
; _LOCalc_ConvertColorFromLong
; _LOCalc_ConvertColorToLong
; _LOCalc_ConvertFromMicrometer
; _LOCalc_ConvertToMicrometer
; _LOCalc_FilterDescriptorCreate
; _LOCalc_FilterDescriptorModify
; _LOCalc_FilterFieldCreate
; _LOCalc_FilterFieldModify
; _LOCalc_FontExists
; _LOCalc_FontsGetNames
; _LOCalc_FormatKeyCreate
; _LOCalc_FormatKeyDelete
; _LOCalc_FormatKeyExists
; _LOCalc_FormatKeyGetStandard
; _LOCalc_FormatKeyGetString
; _LOCalc_FormatKeysGetList
; _LOCalc_PathConvert
; _LOCalc_SearchDescriptorCreate
; _LOCalc_SearchDescriptorModify
; _LOCalc_SearchDescriptorSimilarityModify
; _LOCalc_SortFieldCreate
; _LOCalc_SortFieldModify
; _LOCalc_VersionGet
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LOCalc_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
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
Func _LOCalc_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
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
EndFunc   ;==>_LOCalc_ComError_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_ConvertColorFromLong
; Description ...: Convert Long color code to Hex, RGB, HSB or CMYK.
; Syntax ........: _LOCalc_ConvertColorFromLong([$iHex = Null[, $iRGB = Null[, $iHSB = Null[, $iCMYK = Null]]]])
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
; Related .......: _LOCalc_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_ConvertColorFromLong($iHex = Null, $iRGB = Null, $iHSB = Null, $iCMYK = Null)
	Local $nRed, $nGreen, $nBlue, $nResult, $nMaxRGB, $nMinRGB, $nHue, $nSaturation, $nBrightness, $nCyan, $nMagenta, $nYellow, $nBlack
	Local $dHex
	Local $aiReturn[0]

	If (@NumParams = 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Select
		Case IsInt($iHex) ; Long to Hex
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

		Case IsInt($iHSB) ; Long to HSB
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
EndFunc   ;==>_LOCalc_ConvertColorFromLong

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_ConvertColorToLong
; Description ...: Convert Hex, RGB, HSB or CMYK to Long color code.
; Syntax ........: _LOCalc_ConvertColorToLong([$vVal1 = Null[, $vVal2 = Null[, $vVal3 = Null[, $vVal4 = Null]]]])
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
; Related .......: _LOCalc_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_ConvertColorToLong($vVal1 = Null, $vVal2 = Null, $vVal3 = Null, $vVal4 = Null) ; RGB = Int, CMYK = Int, HSB = String, Hex = String.
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
EndFunc   ;==>_LOCalc_ConvertColorToLong

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_ConvertFromMicrometer
; Description ...: Convert from Micrometer to Inch, Centimeter, Millimeter, or Printer's Points.
; Syntax ........: _LOCalc_ConvertFromMicrometer([$nInchOut = Null[, $nCentimeterOut = Null[, $nMillimeterOut = Null[, $nPointsOut = Null]]]])
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
; Related .......: _LOCalc_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_ConvertFromMicrometer($nInchOut = Null, $nCentimeterOut = Null, $nMillimeterOut = Null, $nPointsOut = Null)
	Local $nReturnValue

	If ($nInchOut <> Null) Then
		If Not IsNumber($nInchOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		$nReturnValue = __LOCalc_UnitConvert($nInchOut, $__LOCONST_CONVERT_UM_INCH)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $nReturnValue)
	EndIf

	If ($nCentimeterOut <> Null) Then
		If Not IsNumber($nCentimeterOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$nReturnValue = __LOCalc_UnitConvert($nCentimeterOut, $__LOCONST_CONVERT_UM_CM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 2, $nReturnValue)
	EndIf

	If ($nMillimeterOut <> Null) Then
		If Not IsNumber($nMillimeterOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$nReturnValue = __LOCalc_UnitConvert($nMillimeterOut, $__LOCONST_CONVERT_UM_MM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Return SetError($__LO_STATUS_SUCCESS, 3, $nReturnValue)
	EndIf

	If ($nPointsOut <> Null) Then
		If Not IsNumber($nPointsOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$nReturnValue = __LOCalc_UnitConvert($nPointsOut, $__LOCONST_CONVERT_UM_PT)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Return SetError($__LO_STATUS_SUCCESS, 4, $nReturnValue)
	EndIf

	Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; NO Unit set.
EndFunc   ;==>_LOCalc_ConvertFromMicrometer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_ConvertToMicrometer
; Description ...: Convert from Inch, Centimeter, Millimeter, or Printer's Points to Micrometer.
; Syntax ........: _LOCalc_ConvertToMicrometer([$nInchIn = Null[, $nCentimeterIn = Null[, $nMillimeterIn = Null[, $nPointsIn = Null]]]])
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
; Remarks .......: To skip a parameter, set it to Null.
;                  If you are converting from Inches, call inches in $nInchIn, if converting from Centimeters, $nInchIn is called with Null, and $nCentimeters is set.
;                  A Micrometer is 1000th of a centimeter, and is used in almost all Libre Office functions that contain a measurement parameter.
; Related .......: _LOCalc_ConvertFromMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_ConvertToMicrometer($nInchIn = Null, $nCentimeterIn = Null, $nMillimeterIn = Null, $nPointsIn = Null)
	Local $nReturnValue

	If ($nInchIn <> Null) Then
		If Not IsNumber($nInchIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		$nReturnValue = __LOCalc_UnitConvert($nInchIn, $__LOCONST_CONVERT_INCH_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $nReturnValue)
	EndIf

	If ($nCentimeterIn <> Null) Then
		If Not IsNumber($nCentimeterIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$nReturnValue = __LOCalc_UnitConvert($nCentimeterIn, $__LOCONST_CONVERT_CM_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 2, $nReturnValue)
	EndIf

	If ($nMillimeterIn <> Null) Then
		If Not IsNumber($nMillimeterIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$nReturnValue = __LOCalc_UnitConvert($nMillimeterIn, $__LOCONST_CONVERT_MM_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Return SetError($__LO_STATUS_SUCCESS, 3, $nReturnValue)
	EndIf

	If ($nPointsIn <> Null) Then
		If Not IsNumber($nPointsIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$nReturnValue = __LOCalc_UnitConvert($nPointsIn, $__LOCONST_CONVERT_PT_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Return SetError($__LO_STATUS_SUCCESS, 4, $nReturnValue)
	EndIf

	Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; NO Unit set.
EndFunc   ;==>_LOCalc_ConvertToMicrometer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FilterDescriptorCreate
; Description ...: Create a Filter Descriptor to use in the Filter function.
; Syntax ........: _LOCalc_FilterDescriptorCreate(ByRef $oRange, $atFilterField[, $bCaseSensitive = False[, $bSkipDupl = False[, $bUseRegExp = False[, $bHeaders = False[, $bCopyOutput = False[, $oCopyOutput = Null[, $bSaveCriteria = True]]]]]]])
; Parameters ....: $oRange              - [in/out] an object. The Range you intend to apply the Filter to. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $atFilterField       - an array of dll structs. A single column Array of Filter Fields previously created by _LOCalc_FilterFieldCreate. Maximum of 8 Fields allowed.
;                  $bCaseSensitive      - [optional] a boolean value. Default is False. If True, the Filtering operation will be case sensitive.
;                  $bSkipDupl           - [optional] a boolean value. Default is False. If True, Duplicate values will be skipped in the list of filtered data.
;                  $bUseRegExp          - [optional] a boolean value. Default is False. If True, the String Value set will be considered as using Regular expressions.
;                  $bHeaders            - [optional] a boolean value. Default is False. If True, the Range contains column headers.
;                  $bCopyOutput         - [optional] a boolean value. Default is False. If True, the filtering results are copied to another location in the Sheet.
;                  $oCopyOutput         - [optional] an object. Default is Null. The location to copy filter data to. If a range is input, the first cell is used. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bSaveCriteria       - [optional] a boolean value. Default is True. If True, the output range remains linked to the source range, allowing for future re-application of the same filter to the range. Source Range must be previously defined as a Database range.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $atFilterField not an Array, or Array contains more than 8 elements.
;                  @Error 1 @Extended 3 Return 0 = $bCaseSensitive not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bSkipDupl not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bUseRegExp not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bHeaders not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bCopyOutput not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $oCopyOutput not an Object and not set to Null.
;                  @Error 1 @Extended 9 Return 0 = $bSaveCriteria not a Boolean.
;                  @Error 1 @Extended 10 Return ? = $atFilterField contains an element that is not an Object. Returning the element number containing the error.
;                  @Error 1 @Extended 11 Return 0 = $bCopyOutput set to True, but $oCopyOutput not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Filter Descriptor Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.table.CellAddress" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Address for Cell or Cell Range called in $oCopyOutput.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully created a Filter descriptor Object, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_FilterDescriptorModify, _LOCalc_FilterFieldCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FilterDescriptorCreate(ByRef $oRange, $atFilterField, $bCaseSensitive = False, $bSkipDupl = False, $bUseRegExp = False, $bHeaders = False, $bCopyOutput = False, $oCopyOutput = Null, $bSaveCriteria = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFilterDesc
	Local $tCellInputAddr, $tCellAddr
	Local Const $__LOC_FILTER_ORIENTATION_ROWS = 1 ; Orientation isn't implemented in L.O. so Rows is the only option.

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($atFilterField) Or (UBound($atFilterField) > 8) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bCaseSensitive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bSkipDupl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bUseRegExp) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bHeaders) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If ($oCopyOutput <> Null) And Not IsObj($oCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If Not IsBool($bSaveCriteria) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	$oFilterDesc = $oRange.createFilterDescriptor(True)
	If Not IsObj($oFilterDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($atFilterField) - 1
		If Not IsObj($atFilterField[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, $i)
	Next

	If ($bCopyOutput = True) Then
		If Not IsObj($oCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tCellInputAddr = $oCopyOutput.RangeAddress()
		If Not IsObj($tCellInputAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$tCellAddr = __LOCalc_CreateStruct("com.sun.star.table.CellAddress")
		If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tCellAddr.Sheet = $tCellInputAddr.Sheet()
		$tCellAddr.Column = $tCellInputAddr.StartColumn()
		$tCellAddr.Row = $tCellInputAddr.StartRow()
	EndIf

	; Orientation is only set to Rows. I tried setting it to columns, but it doesn't work. Seemingly Filtering Columns isn't implemented yet, which is confirmed by a
	; post from 2009 by Villeroy on the OpenOffice Forums inside of a Macro posted.
	; https://forum.openoffice.org/en/forum/viewtopic.php?p=78786&sid=1e046304b59035364caecb0ad0a10327#p78786

	With $oFilterDesc
		.setFilterFields2($atFilterField)
		.IsCaseSensitive = $bCaseSensitive
		.SkipDuplicates = $bSkipDupl
		.UseRegularExpressions = $bUseRegExp
		.ContainsHeader = $bHeaders
		.Orientation = $__LOC_FILTER_ORIENTATION_ROWS
		.CopyOutputData = $bCopyOutput
		.SaveOutputPosition = $bSaveCriteria
	EndWith

	If IsObj($oCopyOutput) Then $oFilterDesc.OutputPosition = $tCellAddr

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFilterDesc)
EndFunc   ;==>_LOCalc_FilterDescriptorCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FilterDescriptorModify
; Description ...: Set or Retrieve Filter Descriptor settings.
; Syntax ........: _LOCalc_FilterDescriptorModify(ByRef $oRange, ByRef $oFilterDesc[, $atFilterField = Null[, $bCaseSensitive = Null[, $bSkipDupl = Null[, $bUseRegExp = Null[, $bHeaders = Null[, $bCopyOutput = Null[, $oCopyOutput = Null[, $bSaveCriteria = Null]]]]]]]])
; Parameters ....: $oRange              - [in/out] an object. The Sheet the Filter Descriptor was Created with, or the Range you intend to apply the Filter to. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oFilterDesc         - [in/out] an object. A Filter Descriptor created by a previous _LOCalc_FilterDescriptorCreate function.
;                  $atFilterField       - [optional] an array of dll structs. Default is Null. A single column Array of Filter Fields previously created by _LOCalc_FilterFieldCreate. Maximum of 8 Fields allowed.
;                  $bCaseSensitive      - [optional] a boolean value. Default is Null. If True, the Filtering operation will be case sensitive.
;                  $bSkipDupl           - [optional] a boolean value. Default is Null. If True, Duplicate values will be skipped in the list of filtered data.
;                  $bUseRegExp          - [optional] a boolean value. Default is Null. If True, the String Value set will be considered as using Regular expressions.
;                  $bHeaders            - [optional] a boolean value. Default is Null. If True, the Range contains column headers.
;                  $bCopyOutput         - [optional] a boolean value. Default is Null. If True, the filtering results are copied to another location in the Sheet.
;                  $oCopyOutput         - [optional] an object. Default is Null. The location to copy filter data to. If a range is input, the first cell is used. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bSaveCriteria       - [optional] a boolean value. Default is Null. If True, the output range remains linked to the source range, allowing for future re-application of the same filter to the range. Source Range must be previously defined as a Database range.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oFilterDesc not an Object.
;                  @Error 1 @Extended 3 Return 0 = $atFilterField not an Array, or Array contains more than 8 elements.
;                  @Error 1 @Extended 4 Return ? = $atFilterField contains an element that is not an Object. Returning the element number containing the error.
;                  @Error 1 @Extended 5 Return 0 = $bCaseSensitive not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bSkipDupl not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bUseRegExp not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bHeaders not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bCopyOutput not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bCopyOutput set to True, but $oCopyOutput not an Object.
;                  @Error 1 @Extended 11 Return 0 = $oCopyOutput not an Object.
;                  @Error 1 @Extended 12 Return 0 = $bSaveCriteria not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.table.CellAddress" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Object for Cell referenced in $oCopyOutput.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Cell Address for Cell or Cell Range called in $oCopyOutput.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Filter Descriptor was successfully modified.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When retrieving the current settings for a filter descriptor, the Return value for $oCopyOutput is a single Cell Object.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_FilterDescriptorCreate, _LOCalc_FilterFieldCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FilterDescriptorModify(ByRef $oRange, ByRef $oFilterDesc, $atFilterField = Null, $bCaseSensitive = Null, $bSkipDupl = Null, $bUseRegExp = Null, $bHeaders = Null, $bCopyOutput = Null, $oCopyOutput = Null, $bSaveCriteria = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avFilter[8]
	Local $tCellInputAddr, $tCellAddr
	Local $oCell

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFilterDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOCalc_VarsAreNull($atFilterField, $bCaseSensitive, $bSkipDupl, $bUseRegExp, $bHeaders, $bCopyOutput, $oCopyOutput, $bSaveCriteria) Then
		$oCell = $oRange.Spreadsheet.getCellByPosition($oFilterDesc.OutputPosition.Column(), $oFilterDesc.OutputPosition.Row())
		If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		__LOCalc_ArrayFill($avFilter, $oFilterDesc.getFilterFields2(), $oFilterDesc.IsCaseSensitive(), $oFilterDesc.SkipDuplicates(), $oFilterDesc.UseRegularExpressions(), _
				$oFilterDesc.ContainsHeader(), $oFilterDesc.CopyOutputData(), $oCell, $oFilterDesc.SaveOutputPosition())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFilter)
	EndIf

	If ($atFilterField <> Null) Then
		If Not IsArray($atFilterField) Or Not (UBound($atFilterField) <= 8) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		For $i = 0 To UBound($atFilterField) - 1
			If Not IsObj($atFilterField[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
		Next

		$oFilterDesc.setFilterFields2($atFilterField)
	EndIf

	If ($bCaseSensitive <> Null) Then
		If Not IsBool($bCaseSensitive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFilterDesc.IsCaseSensitive = $bCaseSensitive
	EndIf

	If ($bSkipDupl <> Null) Then
		If Not IsBool($bSkipDupl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFilterDesc.SkipDuplicates = $bSkipDupl
	EndIf

	If ($bUseRegExp <> Null) Then
		If Not IsBool($bUseRegExp) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFilterDesc.UseRegularExpressions = $bUseRegExp
	EndIf

	If ($bHeaders <> Null) Then
		If Not IsBool($bHeaders) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFilterDesc.ContainsHeader = $bHeaders
	EndIf

	If ($bCopyOutput <> Null) Then
		If Not IsBool($bCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If ($bCopyOutput = True) And Not IsObj($oCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oFilterDesc.CopyOutputData = $bCopyOutput
	EndIf

	If ($oCopyOutput <> Null) Then
		If Not IsObj($oCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tCellInputAddr = $oCopyOutput.RangeAddress()
		If Not IsObj($tCellInputAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$tCellAddr = __LOCalc_CreateStruct("com.sun.star.table.CellAddress")
		If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tCellAddr.Sheet = $tCellInputAddr.Sheet()
		$tCellAddr.Column = $tCellInputAddr.StartColumn()
		$tCellAddr.Row = $tCellInputAddr.StartRow()

		$oFilterDesc.OutputPosition = $tCellAddr
	EndIf

	If ($bSaveCriteria <> Null) Then
		If Not IsBool($bSaveCriteria) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oFilterDesc.SaveOutputPosition = $bSaveCriteria
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_FilterDescriptorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FilterFieldCreate
; Description ...: Create a Filter Field for defining Filter values and settings.
; Syntax ........: _LOCalc_FilterFieldCreate($iColumn[, $bIsNumeric = False[, $nValue = 0[, $sString = ""[, $iCondition = $LOC_FILTER_CONDITION_EMPTY[, $iOperator = $LOC_FILTER_OPERATOR_AND]]]]])
; Parameters ....: $iColumn             - an integer value. The 0 based Column number to perform the filtering operation upon counting from the beginning of the range.
;                  $bIsNumeric          - [optional] a boolean value. Default is False. If True, the filter Value to search for is a number. If False, the filter value to search for is a string.
;                  $nValue              - [optional] a general number value. Default is 0. The numerical Value to filter the Range for. Only valid if $bIsNumeric is set to True. Set to any number to skip, it will not be used unless $bIsNumeric is True.
;                  $sString             - [optional] a string value. Default is "". The string Value to filter the Range for. Only valid if $bIsNumeric is set to False. Set to an empty string to skip, it will not be used unless $bIsNumeric is False.
;                  $iCondition          - [optional] an integer value (0-17). Default is $LOC_FILTER_CONDITION_EMPTY. The comparative condition to test each cell and value by. See Constants $LOC_FILTER_CONDITION_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iOperator           - [optional] an integer value (0,1). Default is $LOC_FILTER_OPERATOR_AND. The connection this filter field has with the previous filter field. See Constants $LOC_FILTER_OPERATOR_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Struct
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iColumn not an Integer.
;                  @Error 1 @Extended 2 Return 0 = $bIsNumeric not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $nValue not a number.
;                  @Error 1 @Extended 4 Return 0 = $sString not a String.
;                  @Error 1 @Extended 5 Return 0 = $iCondition not an Integer, less than 0 or greater than 17. See Constants $LOC_FILTER_CONDITION_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iOperator not an Integer, less than 0 or greater than 1. See Constants $LOC_FILTER_OPERATOR_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.sheet.TableFilterField2" Struct.
;                  --Success--
;                  @Error 0 @Extended 0 Return Struct = Success. Successfully created and returned the Filter Field Structure.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Filter Descriptor can contain up to 8 of these Filter Fields. Once you create the Filter Field Structure, place it in an array before using it to create a Filter descriptor. Place each Filter Field Structure in a separate element of the Array.
; Related .......: _LOCalc_FilterFieldModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FilterFieldCreate($iColumn, $bIsNumeric = False, $nValue = 0, $sString = "", $iCondition = $LOC_FILTER_CONDITION_EMPTY, $iOperator = $LOC_FILTER_OPERATOR_AND)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tFilterField

	If Not IsInt($iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bIsNumeric) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LOCalc_IntIsBetween($iCondition, $LOC_FILTER_CONDITION_EMPTY, $LOC_FILTER_CONDITION_DOES_NOT_END_WITH) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not __LOCalc_IntIsBetween($iOperator, $LOC_FILTER_OPERATOR_AND, $LOC_FILTER_OPERATOR_OR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$tFilterField = __LOCalc_CreateStruct("com.sun.star.sheet.TableFilterField2")
	If Not IsObj($tFilterField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	With $tFilterField
		.Field = $iColumn
		.IsNumeric = $bIsNumeric
		.NumericValue = $nValue
		.StringValue = $sString
		.Operator = $iCondition ; L.O. calls Operator "Condition" in U.I.
		.Connection = $iOperator ; L.O. calls Connection "Operator" in U.I.
	EndWith

	Return SetError($__LO_STATUS_SUCCESS, 0, $tFilterField)
EndFunc   ;==>_LOCalc_FilterFieldCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FilterFieldModify
; Description ...: Set or Retrieve Filter Field structure settings.
; Syntax ........: _LOCalc_FilterFieldModify(ByRef $tFilterField[, $iColumn = Null[, $bIsNumeric = Null[, $nValue = Null[, $sString = Null[, $iCondition = Null[, $iOperator = Null]]]]]])
; Parameters ....: $tFilterField        - [in/out] a dll struct value. A Filter Field from a previous _LOCalc_FilterFieldCreate function call.
;                  $iColumn             - [optional] an integer value. Default is Null. The 0 based Column number to perform the filtering operation upon counting from the beginning of the range.
;                  $bIsNumeric          - [optional] a boolean value. Default is Null. If True, the filter Value to search for is a number. If False, the filter value to search for is a string.
;                  $nValue              - [optional] a general number value. Default is Null. The numerical Value to filter the Range for. Only valid if $bIsNumeric is set to True.
;                  $sString             - [optional] a string value. Default is Null. The string Value to filter the Range for. Only valid if $bIsNumeric is set to False.
;                  $iCondition          - [optional] an integer value (0-17). Default is Null. The comparative condition to test each cell and value by.
;                  $iOperator           - [optional] an integer value (0,1). Default is Null. The connection this filter field has with the previous filter field.
; Return values .: Success: Struct
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $tFilterField not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumn not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $bIsNumeric not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $nValue not a number.
;                  @Error 1 @Extended 5 Return 0 = $sString not a String.
;                  @Error 1 @Extended 6 Return 0 = $iCondition not an Integer, less than 0 or greater than 17. See Constants $LOC_FILTER_CONDITION_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $iOperator not an Integer, less than 0 or greater than 1. See Constants $LOC_FILTER_OPERATOR_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Filter Field Structure was successfully modified.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Filter Descriptor can contain up to 8 of these Filter Fields. Once you create the Filter Field Structure, place it in an array before using it to create a Filter descriptor. Place each Filter Field Structure in a separate element of the Array.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_FilterFieldCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FilterFieldModify(ByRef $tFilterField, $iColumn = Null, $bIsNumeric = Null, $nValue = Null, $sString = Null, $iCondition = Null, $iOperator = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avFilter[6]

	If Not IsObj($tFilterField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOCalc_VarsAreNull($iColumn, $bIsNumeric, $nValue, $sString, $iCondition, $iOperator) Then
		__LOCalc_ArrayFill($avFilter, $tFilterField.Field(), $tFilterField.IsNumeric(), $tFilterField.NumericValue(), $tFilterField.StringValue(), $tFilterField.Operator(), $tFilterField.Connection())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFilter)
	EndIf

	If ($iColumn <> Null) Then
		If Not IsInt($iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tFilterField.Field = $iColumn
	EndIf

	If ($bIsNumeric <> Null) Then
		If Not IsBool($bIsNumeric) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tFilterField.IsNumeric = $bIsNumeric
	EndIf

	If ($nValue <> Null) Then
		If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tFilterField.NumericValue = $nValue
	EndIf

	If ($sString <> Null) Then
		If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tFilterField.StringValue = $sString
	EndIf

	If ($iCondition <> Null) Then ; L.O. calls Operator "Condition" in U.I.
		If Not __LOCalc_IntIsBetween($iCondition, $LOC_FILTER_CONDITION_EMPTY, $LOC_FILTER_CONDITION_DOES_NOT_END_WITH) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tFilterField.Operator = $iCondition
	EndIf

	If ($iOperator <> Null) Then ; L.O. calls Connection "Operator" in U.I.
		If Not __LOCalc_IntIsBetween($iOperator, $LOC_FILTER_OPERATOR_AND, $LOC_FILTER_OPERATOR_OR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tFilterField.Connection = $iOperator
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_FilterFieldModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FontExists
; Description ...: Tests whether a specific font exists by name.
; Syntax ........: _LOCalc_FontExists($sFontName[, $oDoc = Null])
; Parameters ....: $sFontName           - a string value. The Font name to search for.
;                  $oDoc                - [optional] an object. Default is Null. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
; Remarks .......: $oDoc is optional, if not called, a Calc Document is created invisibly to perform the check.
; Related .......: _LOCalc_FontsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FontExists($sFontName, $oDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts, $atProperties[1]
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $oServiceManager, $oDesktop
	Local $bClose = False

	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If Not IsObj($oDoc) Then
		$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
		If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$atProperties[0] = __LOCalc_SetPropertyValue("Hidden", True)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$oDoc = $oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", $iURLFrameCreate, $atProperties)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

		$bClose = True
	EndIf

	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then
		If $bClose Then $oDoc.Close(True)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	For $i = 0 To UBound($atFonts) - 1
		If $atFonts[$i].Name = $sFontName Then
			If $bClose Then $oDoc.Close(True)

			Return SetError($__LO_STATUS_SUCCESS, 0, True)
		EndIf
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If $bClose Then $oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOCalc_FontExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FontsGetNames
; Description ...: Retrieve an array of currently available font names.
; Syntax ........: _LOCalc_FontsGetNames([$oDoc = Null])
; Parameters ....: $oDoc                - [optional] an object. Default is Null. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  @Error 2 @Extended 3 Return 0 = Failed to create a Property Struct.
;                  @Error 2 @Extended 4 Return 0 = Failed to create a new Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Font list.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returns a 4 Column Array, @extended is set to the number of results. See remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oDoc is optional, if not called, a Calc Document is created invisibly to perform the check.
;                  Many fonts will be listed multiple times, this is because of the varying settings for them, such as bold, Italic, etc. Style Name is really a repeat of weight(Bold) and Slant (Italic) settings, but is included for easier processing if required.
;                  From personal tests, Slant only returns 0 or 2.
;                  The returned array will be as follows:
;                  The first column (Array[1][0]) contains the Font Name.
;                  The Second column (Array [1][1] contains the style name (Such as Bold Italic etc.)
;                  The third column (Array[1][2]) contains the Font weight (Bold) See Constants, $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3;
;                  The fourth column (Array[1][3]) contains the font slant (Italic) See constants, $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
; Related .......: _LOCalc_FontExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FontsGetNames($oDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asFonts[0][4]
	Local $atFonts, $atProperties[1]
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $oServiceManager, $oDesktop
	Local $bClose = False

	If Not IsObj($oDoc) Then
		$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
		If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$atProperties[0] = __LOCalc_SetPropertyValue("Hidden", True)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$oDoc = $oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", $iURLFrameCreate, $atProperties)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

		$bClose = True
	EndIf

	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then
		If $bClose Then $oDoc.Close(True)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	ReDim $asFonts[UBound($atFonts)][4]

	For $i = 0 To UBound($atFonts) - 1
		$asFonts[$i][0] = $atFonts[$i].Name()
		$asFonts[$i][1] = $atFonts[$i].StyleName()
		$asFonts[$i][2] = $atFonts[$i].Weight
		$asFonts[$i][3] = $atFonts[$i].Slant() ; only 0 or 2?
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If $bClose Then $oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, UBound($atFonts), $asFonts)
EndFunc   ;==>_LOCalc_FontsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeyCreate
; Description ...: Create a Format Key.
; Syntax ........: _LOCalc_FormatKeyCreate(ByRef $oDoc, $sFormat)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sFormat             - a string value. The format key String to create.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFormat not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to Create or Retrieve the Format key.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Format Key was successfully created, returning Format Key integer.
;                  @Error 0 @Extended 1 Return Integer = Success. Format Key already existed, returning Format Key integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_FormatKeyDelete, _LOCalc_FormatKeyGetStandard
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeyCreate(ByRef $oDoc, $sFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFormatKey
	Local $tLocale
	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tLocale = __LOCalc_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iFormatKey = $oFormats.queryKey($sFormat, $tLocale, False)
	If ($iFormatKey > -1) Then Return SetError($__LO_STATUS_SUCCESS, 1, $iFormatKey) ; Format already existed
	$iFormatKey = $oFormats.addNew($sFormat, $tLocale)
	If ($iFormatKey > -1) Then Return SetError($__LO_STATUS_SUCCESS, 0, $iFormatKey) ; Format created

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Failed to create or retrieve Format
EndFunc   ;==>_LOCalc_FormatKeyCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeyDelete
; Description ...: Delete a User-Created Format Key from a Document.
; Syntax ........: _LOCalc_FormatKeyDelete(ByRef $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iFormatKey          - an integer value. The User-Created format Key to delete.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 3 Return 0 = Format Key called in $iFormatKey not found in Document.
;                  @Error 1 @Extended 4 Return 0 = Format Key called in $iFormatKey not User-Created.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to delete key.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Format Key was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_FormatKeysGetList, _LOCalc_FormatKeyCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeyDelete(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOCalc_FormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Key not found.

	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($oFormats.getbykey($iFormatKey).UserDefined() = False) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Key not User Created.

	$oFormats.removeByKey($iFormatKey)

	Return (_LOCalc_FormatKeyExists($oDoc, $iFormatKey) = False) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0))
EndFunc   ;==>_LOCalc_FormatKeyDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeyExists
; Description ...: Check if a Document contains a certain Format Key.
; Syntax ........: _LOCalc_FormatKeyExists(ByRef $oDoc, $iFormatKey[, $iFormatType = $LOC_FORMAT_KEYS_ALL])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iFormatKey          - an integer value. The Format Key to look for.
;                  $iFormatType         - [optional] an integer value (0-15881). Default is $LOC_FORMAT_KEYS_ALL. The Format Key type to search in. Values can be BitOr'd together. See Constants, $LOC_FORMAT_KEYS_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iFormatType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to obtain Array of Date/Time Formats.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the Format Key exists in document, True is returned, else false.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeyExists(ByRef $oDoc, $iFormatKey, $iFormatType = $LOC_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys[0]
	Local $tLocale

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iFormatType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tLocale = __LOCalc_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aiFormatKeys = $oFormats.queryKeys($iFormatType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aiFormatKeys) - 1
		If ($aiFormatKeys[$i] = $iFormatKey) Then Return SetError($__LO_STATUS_SUCCESS, 0, True) ; Doc does contain format Key
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, False) ; Doc does not contain format Key
EndFunc   ;==>_LOCalc_FormatKeyExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeyGetStandard
; Description ...: Retrieve the Standard Format for a specific Format Key Type.
; Syntax ........: _LOCalc_FormatKeyGetStandard(ByRef $oDoc, $iFormatKeyType)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iFormatKeyType      - an integer value (1-8196). The Format Key type to retrieve the standard Format for. See Constants $LOC_FORMAT_KEYS_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iFormatKeyType not an Integer, less than 1 or greater than 8196. See Constants $LOC_FORMAT_KEYS_* as defined in LibreOfficeCalc_Constants.au3.
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
Func _LOCalc_FormatKeyGetStandard(ByRef $oDoc, $iFormatKeyType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $tLocale
	Local $iStandard

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iFormatKeyType, $LOC_FORMAT_KEYS_DEFINED, $LOC_FORMAT_KEYS_DURATION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tLocale = __LOCalc_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iStandard = $oFormats.getStandardFormat($iFormatKeyType, $tLocale)
	If Not IsInt($iStandard) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iStandard)
EndFunc   ;==>_LOCalc_FormatKeyGetStandard

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeyGetString
; Description ...: Retrieve a Format Key String.
; Syntax ........: _LOCalc_FormatKeyGetString(ByRef $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iFormatKey          - an integer value. The Format Key to retrieve the string for.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iFormatKey not found in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve requested Format Key Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning Format Key's Format String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_FormatKeysGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeyGetString(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormatKey

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOCalc_FormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oFormatKey = $oDoc.getNumberFormats().getByKey($iFormatKey)
	If Not IsObj($oFormatKey) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Key not found.

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFormatKey.FormatString())
EndFunc   ;==>_LOCalc_FormatKeyGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeysGetList
; Description ...: Retrieve an Array of Date/Time Format Keys.
; Syntax ........: _LOCalc_FormatKeysGetList(ByRef $oDoc[, $bIsUser = False[, $bUserOnly = False[, $iFormatKeyType = $LOC_FORMAT_KEYS_ALL]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bIsUser             - [optional] a boolean value. Default is False. If True, Adds a third column to the return Array with a boolean, whether each Key is user-created or not.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True, only user-created Format Keys are returned.
;                  $iFormatKeyType      - [optional] an integer value (0-15881). Default is $LOC_FORMAT_KEYS_ALL. The Format Key type to retrieve an array of. Values can be BitOr'd together. See Constants, $LOC_FORMAT_KEYS_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bIsUser not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bUserOnly not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $iFormatKeyType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve NumberFormats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to obtain Array of Format Keys.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning a 2 or 3 column Array, depending on current $bIsUser setting. See remarks. @Extended is set to the number of Keys returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Column One (Array[0][0]) will contain the Format Key integer,
;                  Column two (Array[0][1]) will contain the Format Key String,
;                  If $bIsUser is set to True, Column Three (Array[0][2]) will contain a Boolean, True if the Format Key is User-created, else false.
; Related .......: _LOCalc_FormatKeyDelete, _LOCalc_FormatKeyGetString, _LOCalc_FormatKeyGetStandard
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeysGetList(ByRef $oDoc, $bIsUser = False, $bUserOnly = False, $iFormatKeyType = $LOC_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys
	Local $avFormats[0][3]
	Local $tLocale
	Local $iColumns = 3, $iCount = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bIsUser) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iColumns = ($bIsUser = True) ? ($iColumns) : (2)

	If Not IsInt($iFormatKeyType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$tLocale = __LOCalc_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oFormats = $oDoc.getNumberFormats()
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
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	If ($bUserOnly = True) Then ReDim $avFormats[$iCount][$iColumns]

	Return SetError($__LO_STATUS_SUCCESS, UBound($avFormats), $avFormats)
EndFunc   ;==>_LOCalc_FormatKeysGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_PathConvert
; Description ...: Converts the input path to or from a LibreOffice URL notation path.
; Syntax ........: _LOCalc_PathConvert($sFilePath[, $iReturnMode = $LOC_PATHCONV_AUTO_RETURN])
; Parameters ....: $sFilePath           - a string value. Full path to convert in String format.
;                  $iReturnMode         - [optional] an integer value (0-2). Default is $__g_iAutoReturn. The type of path format to return. See Constants, $LOC_PATHCONV_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFilePath is not a string
;                  @Error 1 @Extended 2 Return 0 = $iReturnMode not a Integer, less than 0, or greater than 2, see constants, $LOC_PATHCONV_* as defined in LibreOfficeCalc_Constants.au3.
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
Func _LOCalc_PathConvert($sFilePath, $iReturnMode = $LOC_PATHCONV_AUTO_RETURN)
	Local Const $STR_STRIPLEADING = 1
	Local $asURLReplace[9][2] = [["%", "%25"], [" ", "%20"], ["\", "/"], [";", "%3B"], ["#", "%23"], ["^", "%5E"], ["{", "%7B"], ["}", "%7D"], ["`", "%60"]]
	Local $iPathSearch, $iFileSearch, $iPartialPCPath, $iPartialFilePath

	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iReturnMode, $LOC_PATHCONV_AUTO_RETURN, $LOC_PATHCONV_PCPATH_RETURN) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sFilePath = StringStripWS($sFilePath, $STR_STRIPLEADING)

	$iPathSearch = StringRegExp($sFilePath, "[A-Z]\:\\") ; Search For a Computer Path, as in C:\ etc.
	$iPartialPCPath = StringInStr($sFilePath, "\") ; Search for partial computer Path containing a backslash.
	$iFileSearch = StringInStr($sFilePath, "file:///", 0, 1, 1, 9) ; Search for a full Libre path, which begins with File:///
	$iPartialFilePath = StringInStr($sFilePath, "/") ; Search For a Partial Libre path containing forward slash

	If ($iReturnMode = $LOC_PATHCONV_AUTO_RETURN) Then
		If ($iPathSearch > 0) Or ($iPartialPCPath > 0) Then ;  if file path contains partial or full PC path, set to convert to Libre URL.
			$iReturnMode = $LOC_PATHCONV_OFFICE_RETURN

		ElseIf ($iFileSearch > 0) Or ($iPartialFilePath > 0) Then ;  if file path contains partial or full Libre URL, set to convert to PC Path.
			$iReturnMode = $LOC_PATHCONV_PCPATH_RETURN

		Else ; If file path contains neither above. convert to Libre URL
			$iReturnMode = $LOC_PATHCONV_OFFICE_RETURN
		EndIf
	EndIf

	Switch $iReturnMode
		Case $LOC_PATHCONV_OFFICE_RETURN
			If $iFileSearch > 0 Then Return SetError($__LO_STATUS_SUCCESS, 2, $sFilePath)
			If ($iPathSearch > 0) Then $sFilePath = "file:///" & $sFilePath

			For $i = 0 To (UBound($asURLReplace) - 1)
				$sFilePath = StringReplace($sFilePath, $asURLReplace[$i][0], $asURLReplace[$i][1])
				Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
			Next

			Return SetError($__LO_STATUS_SUCCESS, 2, $sFilePath)

		Case $LOC_PATHCONV_PCPATH_RETURN
			If ($iPathSearch > 0) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sFilePath)
			If ($iFileSearch > 0) Then $sFilePath = StringReplace($sFilePath, "file:///", Null)

			For $i = 0 To (UBound($asURLReplace) - 1)
				$sFilePath = StringReplace($sFilePath, $asURLReplace[$i][1], $asURLReplace[$i][0])
				Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
			Next

			Return SetError($__LO_STATUS_SUCCESS, 1, $sFilePath)
	EndSwitch
EndFunc   ;==>_LOCalc_PathConvert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SearchDescriptorCreate
; Description ...: Create a Search Descriptor for searching a document.
; Syntax ........: _LOCalc_SearchDescriptorCreate(ByRef $oRange[, $bBackwards = False[, $bSearchRows = True[, $bMatchCase = False[, $iSearchIn = $LOC_SEARCH_IN_FORMULAS[, $bEntireCell = False[, $bRegExp = False[, $bWildcards = False[, $bStyles = False]]]]]]]])
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bBackwards          - [optional] a boolean value. Default is False. If True, the document is searched backwards.
;                  $bSearchRows         - [optional] a boolean value. Default is True. If True, Search is performed left to right along the rows, else if False, the search is performed top to bottom along the columns.
;                  $bMatchCase          - [optional] a boolean value. Default is False. If True, the case of the letters is important for the Search.
;                  $iSearchIn           - [optional] an integer value. Default is $LOC_SEARCH_IN_FORMULAS. Set the Cell data type to search in. See Constants $LOC_SEARCH_IN_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bEntireCell         - [optional] a boolean value. Default is False. If True, Searches for whole words or cells that are identical to the search text.
;                  $bRegExp             - [optional] a boolean value. Default is False. If True, the search string is evaluated as a regular expression.
;                  $bWildcards          - [optional] a boolean value. Default is False. If True, the search string is considered to contain wildcards (* ?). A Backslash can be used to escape a wildcard.
;                  $bStyles             - [optional] a boolean value. Default is False. If True, the search string is considered a Cell Style name, and the search will return any Cell utilizing the specified name.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bBackwards not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bSearchRows not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bMatchCase not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iSearchIn not an Integer, less than 0 or greater than 2. See Constants $LOC_SEARCH_IN_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bEntireCell not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bRegExp not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bWildcards not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bStyles not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = Both $bRegExp and $bWildcards are set to True, only one can be True at one time.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Search Descriptor.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returns a Search Descriptor Object for setting Search options.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The returned Search Descriptor is only good for the Document that contained the Range it was created by, it WILL NOT work for other Documents.
; Related .......: _LOCalc_SearchDescriptorModify, _LOCalc_SearchDescriptorSimilarityModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SearchDescriptorCreate(ByRef $oRange, $bBackwards = False, $bSearchRows = True, $bMatchCase = False, $iSearchIn = $LOC_SEARCH_IN_FORMULAS, $bEntireCell = False, $bRegExp = False, $bWildcards = False, $bStyles = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSrchDescript

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bBackwards) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bSearchRows) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bMatchCase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LOCalc_IntIsBetween($iSearchIn, $LOC_SEARCH_IN_FORMULAS, $LOC_SEARCH_IN_COMMENTS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bEntireCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bRegExp) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not IsBool($bWildcards) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If Not IsBool($bStyles) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If ($bWildcards = True) And ($bRegExp = True) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

	$oSrchDescript = $oRange.createSearchDescriptor()
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	With $oSrchDescript
		.SearchBackwards = $bBackwards
		.SearchByRow = $bSearchRows
		.SearchCaseSensitive = $bMatchCase
		.SearchType = $iSearchIn
		.SearchWords = $bEntireCell
		.SearchWildcard = $bWildcards
		; Regular Expression setting MUST be after Wildcards, setting Wildcards to False, even when it is already set to False, changes RegExp to False no matter what.
		; -- Slated to be fixed L.O. 24.8.0
		.SearchRegularExpression = $bRegExp
		.SearchStyles = $bStyles
	EndWith

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSrchDescript)
EndFunc   ;==>_LOCalc_SearchDescriptorCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SearchDescriptorModify
; Description ...: Modify Search Descriptor settings of an existing Search Descriptor Object.
; Syntax ........: _LOCalc_SearchDescriptorModify(ByRef $oSrchDescript[, $bBackwards = Null[, $bSearchRows = Null[, $bMatchCase = Null[, $iSearchIn = Null[, $bEntireCell = Null[, $bRegExp = Null[, $bWildcards = Null[, $bStyles = Null]]]]]]]])
; Parameters ....: $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $bBackwards          - [optional] a boolean value. Default is Null. If True, the document is searched backwards.
;                  $bSearchRows         - [optional] a boolean value. Default is Null. If True, Search is performed left to right along the rows, else if False, the search is performed top to bottom along the columns.
;                  $bMatchCase          - [optional] a boolean value. Default is Null. If True, the case of the letters is important for the Search.
;                  $iSearchIn           - [optional] an integer value. Default is Null. Set the Cell data type to search in. See Constants $LOC_SEARCH_IN_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bEntireCell         - [optional] a boolean value. Default is Null. If True, Searches for whole words or cells that are identical to the search text.
;                  $bRegExp             - [optional] a boolean value. Default is Null. If True, the search string is evaluated as a regular expression.
;                  $bWildcards          - [optional] a boolean value. Default is Null. If True, the search string is considered to contain wildcards (* ?). A Backslash can be used to escape a wildcard.
;                  $bStyles             - [optional] a boolean value. Default is Null. If True, the search string is considered a Cell Style name, and the search will return any Cell utilizing the specified name.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript Object not a Search Descriptor Object.
;                  @Error 1 @Extended 3 Return 0 = $bBackwards not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bSearchRows not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bMatchCase not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iSearchIn not an Integer, less than 0 or greater than 2. See Constants $LOC_SEARCH_IN_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bEntireCell not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bRegExp not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bWildcards not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bStyles not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Returns 1 after directly modifying Search Descriptor Object.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When setting $bRegExp or $bWildcards to True, if any of following three are set to True, they will be set to False: $bSimilarity(From the Similarity function), $bRegExp or $bWildcards.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_SearchDescriptorCreate, _LOCalc_SearchDescriptorSimilarityModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SearchDescriptorModify(ByRef $oSrchDescript, $bBackwards = Null, $bSearchRows = Null, $bMatchCase = Null, $iSearchIn = Null, $bEntireCell = Null, $bRegExp = Null, $bWildcards = Null, $bStyles = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSrchDescript[8]

	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOCalc_VarsAreNull($bBackwards, $bSearchRows, $bMatchCase, $iSearchIn, $bEntireCell, $bRegExp, $bWildcards, $bStyles) Then
		__LOCalc_ArrayFill($avSrchDescript, $oSrchDescript.SearchBackwards(), $oSrchDescript.SearchByRow(), $oSrchDescript.SearchCaseSensitive(), _
				$oSrchDescript.SearchType(), $oSrchDescript.SearchWords(), $oSrchDescript.SearchRegularExpression(), $oSrchDescript.SearchWildcard(), _
				$oSrchDescript.SearchStyles())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSrchDescript)
	EndIf

	If ($bBackwards <> Null) Then
		If Not IsBool($bBackwards) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oSrchDescript.SearchBackwards = $bBackwards
	EndIf

	If ($bSearchRows <> Null) Then
		If Not IsBool($bSearchRows) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oSrchDescript.SearchByRow = $bSearchRows
	EndIf

	If ($bMatchCase <> Null) Then
		If Not IsBool($bMatchCase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oSrchDescript.SearchCaseSensitive = $bMatchCase
	EndIf

	If ($iSearchIn <> Null) Then
		If Not __LOCalc_IntIsBetween($iSearchIn, $LOC_SEARCH_IN_FORMULAS, $LOC_SEARCH_IN_COMMENTS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oSrchDescript.SearchType = $iSearchIn
	EndIf

	If ($bEntireCell <> Null) Then
		If Not IsBool($bEntireCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oSrchDescript.SearchWords = $bEntireCell
	EndIf

	If ($bWildcards <> Null) Then
		If Not IsBool($bWildcards) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		If ($bWildcards = True) And ($oSrchDescript.SearchSimilarity = True) Then $oSrchDescript.SearchSimilarity = False
		If ($bWildcards = True) And ($oSrchDescript.SearchRegularExpression = True) Then $oSrchDescript.SearchRegularExpression = False
		$oSrchDescript.SearchWildcard = $bWildcards
	EndIf
	; Regular Expression setting MUST be after Wildcards, setting Wildcards to False, even when it is already set to False, changes RegExp to False no matter what.
	; -- Slated to be fixed L.O. 24.8.0
	If ($bRegExp <> Null) Then
		If Not IsBool($bRegExp) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		If ($bRegExp = True) And ($oSrchDescript.SearchSimilarity = True) Then $oSrchDescript.SearchSimilarity = False
		$oSrchDescript.SearchRegularExpression = $bRegExp
	EndIf

	If ($bStyles <> Null) Then
		If Not IsBool($bStyles) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oSrchDescript.SearchStyles = $bStyles
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SearchDescriptorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SearchDescriptorSimilarityModify
; Description ...: Modify Similarity Search Settings for an existing Search Descriptor Object.
; Syntax ........: _LOCalc_SearchDescriptorSimilarityModify(ByRef $oSrchDescript[, $bSimilarity = Null[, $bCombine = Null[, $iRemove = Null[, $iAdd = Null[, $iExchange = Null]]]]])
; Parameters ....: $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $bSimilarity         - [optional] a boolean value. Default is Null. If True, a "similarity search" is performed.
;                  $bCombine            - [optional] a boolean value. Default is Null. If True, all similarity rules ($iRemove, $iAdd, and $iExchange) are applied together.
;                  $iRemove             - [optional] an integer value. Default is Null. Specifies the number of characters that may be ignored to match the search pattern.
;                  $iAdd                - [optional] an integer value. Default is Null. Specifies the number of characters that must be added to match the search pattern.
;                  $iExchange           - [optional] an integer value. Default is Null. Specifies the number of characters that must be replaced to match the search pattern.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript Object not a Search Descriptor Object.
;                  @Error 1 @Extended 3 Return 0 = $bSimilarity not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bCombine not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iRemove, $iAdd, or $iExchange set to a value, but $bSimilarity not set to True.
;                  @Error 1 @Extended 6 Return 0 = $iRemove not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iAdd not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $iExchange not an Integer.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Returns 1 after directly modifying Search Descriptor Object.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  If $bSimilarity is set to True while Regular Expression, or Wildcards setting is set to True, those settings will be set to False.
; Related .......: _LOCalc_SearchDescriptorCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SearchDescriptorSimilarityModify(ByRef $oSrchDescript, $bSimilarity = Null, $bCombine = Null, $iRemove = Null, $iAdd = Null, $iExchange = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSrchDescript[5]

	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOCalc_VarsAreNull($bSimilarity, $bCombine, $iRemove, $iAdd, $iExchange) Then
		__LOCalc_ArrayFill($avSrchDescript, $oSrchDescript.SearchSimilarity(), $oSrchDescript.SearchSimilarityRelax(), _
				$oSrchDescript.SearchSimilarityRemove(), $oSrchDescript.SearchSimilarityAdd(), $oSrchDescript.SearchSimilarityExchange())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSrchDescript)
	EndIf

	If ($bSimilarity <> Null) Then
		If Not IsBool($bSimilarity) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If ($bSimilarity = True) And ($oSrchDescript.SearchRegularExpression = True) Then $oSrchDescript.SearchRegularExpression = False
		If ($bSimilarity = True) And ($oSrchDescript.SearchWildcard = True) Then $oSrchDescript.SearchWildcard = False
		$oSrchDescript.SearchSimilarity = $bSimilarity
	EndIf

	If ($bCombine <> Null) Then
		If Not IsBool($bCombine) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oSrchDescript.SearchSimilarityRelax = $bCombine
	EndIf

	If Not __LOCalc_VarsAreNull($iRemove, $iAdd, $iExchange) Then
		If ($oSrchDescript.SearchSimilarity() = False) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		If ($iRemove <> Null) Then
			If Not IsInt($iRemove) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oSrchDescript.SearchSimilarityRemove = $iRemove
		EndIf

		If ($iAdd <> Null) Then
			If Not IsInt($iAdd) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oSrchDescript.SearchSimilarityAdd = $iAdd
		EndIf

		If ($iExchange <> Null) Then
			If Not IsInt($iExchange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oSrchDescript.SearchSimilarityExchange = $iExchange
		EndIf
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SearchDescriptorSimilarityModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SortFieldCreate
; Description ...: Create a Sort Field for sorting a Range of data with.
; Syntax ........: _LOCalc_SortFieldCreate($iIndex[, $iDataType = $LOC_SORT_DATA_TYPE_AUTO[, $bAscending = True[, $bCaseSensitive = False]]])
; Parameters ....: $iIndex              - an integer value. The Column or Row to perform the sort upon. 0 Based. 0 is the first Column/Row in the Cell Range.
;                  $iDataType           - [optional] an integer value (0-2). Default is $LOC_SORT_DATA_TYPE_AUTO. The type of data that will be sorted. See Constants $LOC_SORT_DATA_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;                  $bAscending          - [optional] a boolean value. Default is True. If True, data will be sorted into ascending order.
;                  $bCaseSensitive      - [optional] a boolean value. Default is False. If True, sort will be case sensitive.
; Return values .: Success: Struct
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iIndex not an Integer, or less than 0.
;                  @Error 1 @Extended 2 Return 0 = $iDataType not an Integer, less than 0 or greater than 2. See Constants $LOC_SORT_DATA_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;                  @Error 1 @Extended 3 Return 0 = $bAscending not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bCaseSensitive not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.table.TableSortField" Struct.
;                  --Success--
;                  @Error 0 @Extended 0 Return Struct = Success. Successfully created and returned a Sort Field Struct.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SortFieldCreate($iIndex, $iDataType = $LOC_SORT_DATA_TYPE_AUTO, $bAscending = True, $bCaseSensitive = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tSortField

	If Not __LOCalc_IntIsBetween($iIndex, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iDataType, $LOC_SORT_DATA_TYPE_AUTO, $LOC_SORT_DATA_TYPE_ALPHANUMERIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAscending) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bCaseSensitive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$tSortField = __LOCalc_CreateStruct("com.sun.star.table.TableSortField")
	If Not IsObj($tSortField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	With $tSortField
		.Field = $iIndex
		.FieldType = $iDataType
		.IsAscending = $bAscending
		.IsCaseSensitive = $bCaseSensitive
	EndWith

	Return SetError($__LO_STATUS_SUCCESS, 0, $tSortField)
EndFunc   ;==>_LOCalc_SortFieldCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SortFieldModify
; Description ...: Modify or retrieve the settings for a Sort Field previously created by _LOCalc_SortFieldCreate.
; Syntax ........: _LOCalc_SortFieldModify(ByRef $tSortField[, $iIndex = Null[, $iDataType = Null[, $bAscending = Null[, $bCaseSensitive = Null]]]])
; Parameters ....: $tSortField          - [in/out] a dll struct value. A Sort Field Struct created by a previous _LOCalc_SortFieldCreate function.
;                  $iIndex              - [optional] an integer value. Default is Null. The Column or Row to perform the sort upon. 0 Based. 0 is the first Column/Row in the Cell Range.
;                  $iDataType           - [optional] an integer value. Default is Null. The type of data that will be sorted. See Constants $LOC_SORT_DATA_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;                  $bAscending          - [optional] a boolean value. Default is Null. If True, data will be sorted into ascending order.
;                  $bCaseSensitive      - [optional] a boolean value. Default is Null. If True, sort will be case sensitive.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $tSortField not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iIndex not an Integer, or less than 0.
;                  @Error 1 @Extended 3 Return 0 = $iDataType not an Integer, less than 0 or greater than 2. See Constants $LOC_SORT_DATA_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;                  @Error 1 @Extended 4 Return 0 = $bAscending not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bCaseSensitive not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SortFieldModify(ByRef $tSortField, $iIndex = Null, $iDataType = Null, $bAscending = Null, $bCaseSensitive = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSort[4]

	If Not IsObj($tSortField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOCalc_VarsAreNull($iIndex, $iDataType, $bAscending, $bCaseSensitive) Then
		__LOCalc_ArrayFill($avSort, $tSortField.Field(), $tSortField.FieldType(), $tSortField.IsAscending(), $tSortField.IsCaseSensitive())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSort)
	EndIf

	If ($iIndex <> Null) Then
		If Not __LOCalc_IntIsBetween($iIndex, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tSortField.Field = $iIndex
	EndIf

	If ($iDataType <> Null) Then
		If Not __LOCalc_IntIsBetween($iDataType, $LOC_SORT_DATA_TYPE_AUTO, $LOC_SORT_DATA_TYPE_ALPHANUMERIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tSortField.FieldType = $iDataType
	EndIf

	If ($bAscending <> Null) Then
		If Not IsBool($bAscending) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tSortField.IsAscending = $bAscending
	EndIf

	If ($bCaseSensitive <> Null) Then
		If Not IsBool($bCaseSensitive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tSortField.IsCaseSensitive = $bCaseSensitive
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SortFieldModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_VersionGet
; Description ...: Retrieve the current Office version.
; Syntax ........: _LOCalc_VersionGet([$bSimpleVersion = False[, $bReturnName = False]])
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
Func _LOCalc_VersionGet($bSimpleVersion = False, $bReturnName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
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

	$aParamArray[0] = __LOCalc_SetPropertyValue("nodepath", "/org.openoffice.Setup/Product")
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSettings = $oConfigProvider.createInstanceWithArguments($sAccess, $aParamArray)

	$sVersionName = $oSettings.getByName("ooName")

	$sVersion = ($bSimpleVersion) ? ($oSettings.getByName("ooSetupVersion")) : ($oSettings.getByName("ooSetupVersionAboutBox"))

	$sReturn = ($bReturnName) ? ($sVersionName & " " & $sVersion) : ($sVersion)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sReturn)
EndFunc   ;==>_LOCalc_VersionGet
