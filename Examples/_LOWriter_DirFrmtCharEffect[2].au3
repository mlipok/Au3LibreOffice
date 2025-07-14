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

	; Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly.")
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

	; Set the selected text's Font effects to shadow = True.
	_LOWriter_DirFrmtCharEffect($oViewCursor, Null, Null, Null, Null, True)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOWriter_DirFrmtCharEffect($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The selected text's current Font Effects settings are as follows: " & @CRLF & _
			"Relief style (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"Case style (See UDF Constants): " & $avSettings[1] & @CRLF & _
			"Are the words hidden? True/False: " & $avSettings[2] & @CRLF & _
			"Are the words outlined? True/False: " & $avSettings[3] & @CRLF & _
			"Do the words have a shadow? True/False: " & $avSettings[4] & @CRLF & @CRLF & _
			"I will now set shadow to false, and Outline to True.")

	; Set the selected text's Font effects to shadow = False and Outline = True.
	_LOWriter_DirFrmtCharEffect($oViewCursor, Null, Null, Null, True, False)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOWriter_DirFrmtCharEffect($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The selected text's current Font Effects settings are as follows: " & @CRLF & _
			"Relief style (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"Case style (See UDF Constants): " & $avSettings[1] & @CRLF & _
			"Are the words hidden? True/False: " & $avSettings[2] & @CRLF & _
			"Are the words outlined? True/False: " & $avSettings[3] & @CRLF & _
			"Do the words have a shadow? True/False: " & $avSettings[4] & @CRLF & @CRLF & _
			"I will next set Outline to false, and set Hidden to true.")

	; Set the selected text's Font effects Outline to False, and Hidden to true.
	_LOWriter_DirFrmtCharEffect($oViewCursor, Null, Null, True, False, Null)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOWriter_DirFrmtCharEffect($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The selected text's current Font Effects settings are as follows: " & @CRLF & _
			"Relief style (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"Case style (See UDF Constants): " & $avSettings[1] & @CRLF & _
			"Are the words hidden? True/False: " & $avSettings[2] & @CRLF & _
			"Are the words outlined? True/False: " & $avSettings[3] & @CRLF & _
			"Do the words have a shadow? True/False: " & $avSettings[4] & @CRLF & @CRLF & _
			"Press ok to remove direct formatting.")

	; Remove Direct formatting.
	_LOWriter_DirFrmtCharEffect($oViewCursor, Default, Default, Default, Default, Default)
	If @error Then _ERROR($oDoc, "Failed to clear the selected text's direct formatting settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
