#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOCalc_PageStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	; Set Page style Sheet Printing Scale settings to: Scale mode = $LOC_SCALE_FIT_WIDTH_HEIGHT, Width = 3 Pages, Height = 8 Pages
	_LOCalc_PageStyleSheetScale($oPageStyle, $LOC_SCALE_FIT_WIDTH_HEIGHT, 3, 8)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOCalc_PageStyleSheetScale($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Because I know I set the Scale Mode to $LOC_SCALE_FIT_WIDTH_HEIGHT, I know the return will have 3 elements, otherwise it would only contain 2." & @CRLF & _
			"The Page Style's current Sheet Scale settings are as follows: " & @CRLF & _
			"The Scaling Mode is (See UDF Constants): " & $avPageStyleSettings[0] & @CRLF & _
			"The Width set for mode $LOC_SCALE_FIT_WIDTH_HEIGHT, is: " & $avPageStyleSettings[1] & @CRLF & _
			"The Height for mode $LOC_SCALE_FIT_WIDTH_HEIGHT, is: " & $avPageStyleSettings[2])

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
