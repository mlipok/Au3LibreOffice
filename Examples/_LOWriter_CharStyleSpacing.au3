#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oCharStyle
	Local $avCharStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text before I modify the Character style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a Character style.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor right 13 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 13)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select the word "Demonstrate".
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 11, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Character style to "Example" Character style.
	_LOWriter_CharStyleSet($oDoc, $oViewCursor, "Example")
	If @error Then _ERROR($oDoc, "Failed to set the Character style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the "Example" object.
	$oCharStyle = _LOWriter_CharStyleGetObj($oDoc, "Example")
	If @error Then _ERROR($oDoc, "Failed to retrieve Character style object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set "Example" Character style Kerning to auto
	_LOWriter_CharStyleSpacing($oCharStyle, True)
	If @error Then _ERROR($oDoc, "Failed to set the Character style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avCharStyleSettings = _LOWriter_CharStyleSpacing($oCharStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Character style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Character style's current Kerning settings are as follows: " & @CRLF & _
			"Automatically adjust kerning? True/False: " & $avCharStyleSettings[0] & @CRLF & _
			"Kerning value, in Printer's points: " & $avCharStyleSettings[1] & @CRLF & @CRLF & _
			"I will now set a custom kerning value.")

	; Set "Example" Character style Kerning to 1.5
	_LOWriter_CharStyleSpacing($oCharStyle, False, 1.5)
	If @error Then _ERROR($oDoc, "Failed to set the Character style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avCharStyleSettings = _LOWriter_CharStyleSpacing($oCharStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Character style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Character style's current Kerning settings are as follows: " & @CRLF & _
			"Automatically adjust kerning? True/False: " & $avCharStyleSettings[0] & @CRLF & _
			"Kerning value, in Printer's points: " & $avCharStyleSettings[1])

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
