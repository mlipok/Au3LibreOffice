#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell
	Local $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell B2
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Insert some text into Cell B2
	_LOCalc_CellString($oCell, "Some Text in a long sentence")
	If @error Then _ERROR($oDoc, "Failed to set Cell Text. Error:" & @error & " Extended:" & @extended)

	; Set the Cell's underline settings to Words only = True, Underline style $LOC_UNDERLINE_BOLD_DASH_DOT, Underline has Color = True, and Color to $LOC_COLOR_BROWN
	_LOCalc_CellUnderLine($oCell, True, $LOC_UNDERLINE_BOLD_DASH_DOT, True, $LOC_COLOR_BROWN)
	If @error Then _ERROR($oDoc, "Failed to set the Cell's settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOCalc_CellUnderLine($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Cell's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Cell's current underline settings are as follows: " & @CRLF & _
			"Underline words only? True/False: " & $avSettings[0] & @CRLF & _
			"Underline style (See UDF constants): " & $avSettings[1] & @CRLF & _
			"Underline has color? True/False: " & $avSettings[2] & @CRLF & _
			"Underline color, in long color format: " & $avSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
