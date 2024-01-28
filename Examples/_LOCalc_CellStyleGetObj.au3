#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oCellStyle

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now retrieve the default Cell style object, and modify some of its settings.")

	; Retrieve the "Default" Cell Style object.
	$oCellStyle = _LOCalc_CellStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell style object. Error:" & @error & " Extended:" & @extended)

	; Set "Default" Cell Style background color to Red.
	_LOCalc_CellStyleBackColor($oCellStyle, $LOC_COLOR_RED)
	If @error Then _ERROR($oDoc, "Failed to set the Cell style settings. Error:" & @error & " Extended:" & @extended)

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
