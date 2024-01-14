#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oParStyle
	Local $avParStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I modify the Default Paragraph style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style." & @CR & "Next Line" & @CR & "Next Line" & @LF)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Retrieve the "Default Paragraph Style" object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Default Paragraph Style")
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended)

	; Set "Default Paragraph Style" Text Flow settings to, Split par, false, Keep Together, false, Orphans 3, Widows 0(off)
	_LOWriter_ParStyleTxtFlowOpt($oParStyle, False, False, 3, 0)
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStyleTxtFlowOpt($oParStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's current Text Flow settings are as follows: " & @CRLF & _
			"Allow the paragraph to be split between pages? True/False: " & $avParStyleSettings[0] & @CRLF & _
			"Allow a Break between this and the next paragraph? True/False: " & $avParStyleSettings[1] & @CRLF & _
			"How many orphans are allowed? (0 means orphan control is off): " & $avParStyleSettings[2] & @CRLF & _
			"How many widows are allowed? (0 means widow control is off): " & $avParStyleSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
