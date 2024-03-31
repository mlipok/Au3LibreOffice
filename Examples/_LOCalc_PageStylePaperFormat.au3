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

	; Set Width to $LOC_PAPER_WIDTH_10ENVELOPE, and Height to $LOC_PAPER_HEIGHT_TABLOID, Landscape to false
	_LOCalc_PageStylePaperFormat($oPageStyle, $LOC_PAPER_WIDTH_10ENVELOPE, $LOC_PAPER_HEIGHT_TABLOID, False)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avPageStyleSettings = _LOCalc_PageStylePaperFormat($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Paper Format settings are as follows: " & @CRLF & _
			"The paper width, in Micrometers, is: " & $avPageStyleSettings[0] & @CRLF & _
			"The paper Height in Micrometers, is: " & $avPageStyleSettings[1] & @CRLF & _
			"Is the Page set to Landscape? True/False: " & $avPageStyleSettings[2])

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
