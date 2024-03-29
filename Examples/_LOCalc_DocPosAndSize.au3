#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $aiReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Set the X coordinate to 50, Y coordinate to 150, Width to 500, Height to 600
	_LOCalc_DocPosAndSize($oDoc, 50, 150, 500, 600)
	If @error Then _ERROR($oDoc, "Failed to set document settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve current document coordinates. Return will be an array in order of function parameters.
	$aiReturn = _LOCalc_DocPosAndSize($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve document position. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The document's current position and size is as follows: " & @CRLF & _
			"X Coordinate = " & $aiReturn[0] & @CRLF & _
			"Y Coordinate = " & $aiReturn[1] & @CRLF & _
			"The document's width, in pixels, is: " & $aiReturn[2] & @CRLF & _
			"The document's height, in pixels, is: " & $aiReturn[3])

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
