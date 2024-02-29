#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iMicrometers, $iMicrometers2
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOCalc_PageStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	; Turn Header on.
	_LOCalc_PageStyleHeader($oPageStyle, True)
	If @error Then _ERROR($oDoc, "Failed to turn Page Style headers on. Error:" & @error & " Extended:" & @extended)

	; Set Page style Header Border Width settings to: $LOC_BORDERWIDTH_MEDIUM on all four sides.
	_LOCalc_PageStyleHeaderBorderWidth($oPageStyle, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Convert 1/8" to Micrometers
	$iMicrometers = _LOCalc_ConvertToMicrometer(.125)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Convert 1/4" to Micrometers
	$iMicrometers2 = _LOCalc_ConvertToMicrometer(.25)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Set Page style Header Border Padding Width settings to: 1/8" on all sides, and then 1/4" on the bottom.
	_LOCalc_PageStyleHeaderBorderPadding($oPageStyle, $iMicrometers, Null, $iMicrometers2)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOCalc_PageStyleHeaderBorderPadding($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Header Border Padding Width settings are as follows: " & @CRLF & _
			"The ""All"" Border Padding Width is, in Micrometers: " & $avPageStyleSettings[0] & @CRLF & _
			"The Top Border Padding Width is, in Micrometers: " & $avPageStyleSettings[1] & @CRLF & _
			"The Bottom Border Padding Width is, in Micrometers: " & $avPageStyleSettings[2] & @CRLF & _
			"The Left Border Padding Width is, in Micrometers: " & $avPageStyleSettings[3] & @CRLF & _
			"The Right Border Padding Width is, in Micrometers: " & $avPageStyleSettings[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
