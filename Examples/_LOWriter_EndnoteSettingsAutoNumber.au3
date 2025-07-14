#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Endnote at the end of this line. ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Endnote at the ViewCursor
	_LOWriter_EndnoteInsert($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert a Endnote. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Document's Endnote settings to: Number format = $LOW_NUM_STYLE_ROMAN_UPPER, Start Numbering at 5, Before the Endnote label, Place
	; a Emdash and after the Endnote label place a Colon
	_LOWriter_EndnoteSettingsAutoNumber($oDoc, $LOW_NUM_STYLE_ROMAN_UPPER, 5, "—", ":")
	If @error Then _ERROR($oDoc, "Failed to modify Endnote settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Endnote settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_EndnoteSettingsAutoNumber($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Endnote settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Document's current Endnote Auto Numbering settings are as follows: " & @CRLF & _
			"The Auto Numbering Number Style used for Endnotes is, (see UDF Constants): " & $avSettings[0] & @CRLF & _
			"The number to start Endnote AutoNumbering at is: " & $avSettings[1] & @CRLF & _
			"The string before the Endnote label is: " & $avSettings[2] & @CRLF & _
			"The string after the Endnote label is: " & $avSettings[3])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
