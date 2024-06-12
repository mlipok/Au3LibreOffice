#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $bResult1, $bResult2

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Test for a font called "Times New Roman"
	$bResult1 = _LOCalc_FontExists($oDoc, "Times New Roman")
	If @error Then _ERROR($oDoc, "Failed to check for font name existing in document. Error:" & @error & " Extended:" & @extended)

	; Test for a font called "Fake Font"
	$bResult2 = _LOCalc_FontExists($oDoc, "Fake Font")
	If @error Then _ERROR($oDoc, "Failed to check for font name existing in document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is the font called ""Times New Roman"" available? True/False: " & $bResult1 & @CRLF & @CRLF & _
			"Is the font called ""Fake Font"" available? True/False: " & $bResult2)

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
