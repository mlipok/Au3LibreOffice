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

	; Convert 1" to Micrometers
	$iMicrometers = _LOCalc_ConvertToMicrometer(1)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Convert 1/2" to Micrometers
	$iMicrometers2 = _LOCalc_ConvertToMicrometer(.5)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Set Left and Right margins to 1", Top and Bottom Margins to 1/2"
	_LOCalc_PageStyleMargins($oPageStyle, $iMicrometers, $iMicrometers, $iMicrometers2, $iMicrometers2)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avPageStyleSettings = _LOCalc_PageStyleMargins($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Margin settings are as follows: " & @CRLF & _
			"The Left page margin, in Micrometers, is: " & $avPageStyleSettings[0] & @CRLF & _
			"The Right page margin, in Micrometers, is: " & $avPageStyleSettings[1] & @CRLF & _
			"The Top page margin, in Micrometers, is: " & $avPageStyleSettings[2] & @CRLF & _
			"The Bottom page margin, in Micrometers, is: " & $avPageStyleSettings[3])

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
