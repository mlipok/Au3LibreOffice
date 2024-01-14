#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oNumStyle
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a new Numbering Style named "Test Style"
	_LOWriter_NumStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to create a Numbering Style. Error:" & @error & " Extended:" & @extended)

	; See if a Numbering Style called "Test Style" exists.
	$bReturn = _LOWriter_NumStyleExists($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to query for Numbering Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Numbering style called ""Test Style"" exist for this document? True/False: " & $bReturn & @CRLF & @CRLF & _
			"Press Ok to retrieve the Numbering Style's Object and then delete the Numbering Style.")

	; Retrieve the "Test Style" Numbering Style object.
	$oNumStyle = _LOWriter_NumStyleGetObj($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to retrieve Numbering style object. Error:" & @error & " Extended:" & @extended)

	; Delete the newly created Numbering Style.
	_LOWriter_NumStyleDelete($oDoc, $oNumStyle)
	If @error Then _ERROR($oDoc, "Failed to delete a Numbering Style. Error:" & @error & " Extended:" & @extended)

	; See if a Numbering Style called "Test Style" exists.
	$bReturn = _LOWriter_NumStyleExists($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to query for Numbering Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Numbering style called ""Test Style"" exist for this document? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
