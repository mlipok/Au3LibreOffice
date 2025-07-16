#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oTextCursor
	Local $aoPars[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "First Paragraph" & @CR & "Second Paragraph" & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 3 rows, 2 columns
	$oTable = _LOWriter_TableCreate($oDoc, 3, 2)
	If @error Then _ERROR($oDoc, "Failed to create a Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert the Table
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR($oDoc, "Failed to insert the table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a list of Paragraph Objects
	$aoPars = _LOWriter_ParObjCreateList($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of paragraphs. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select the second Paragraph
	_LOWriter_ParObjSelect($oDoc, $aoPars[1])
	If @error Then _ERROR($oDoc, "Failed to select the paragraph. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have selected the second paragraph. I could go forward and use _LOWriter_ParObjCopy to copy it, or I could " & _
			"apply direct formatting if I wanted to. But I will now demonstrate selecting a Table using the object returned by creating it.")

	; Select the Table
	_LOWriter_ParObjSelect($oDoc, $oTable)
	If @error Then _ERROR($oDoc, "Failed to select the table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have selected the Table. I could go forward and use _LOWriter_ParObjCopy to copy it, etc." & _
			" I could also use ViewCursor moves, such as GoTo_Start to locate the Viewcursor at the beginning of the Table. But " & _
			"I will now demonstrate selecting the same content a TextCursor has selected by inputting the TextCursor Object.")
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START, 1, False)

	; Create a TextCursor.
	$oTextCursor = _LOWriter_DocCreateTextCursor($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to create a text cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the TextCursor right 1 word.
	_LOWriter_CursorMove($oTextCursor, $LOW_TEXTCUR_GOTO_NEXT_WORD, 1, False)
	If @error Then _ERROR($oDoc, "Failed to move a text cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the TextCursor right 1 word and select it.
	_LOWriter_CursorMove($oTextCursor, $LOW_TEXTCUR_GOTO_END_OF_WORD, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move a text cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The text cursor currently has data selected, but you can't see it because only selections made with the ViewCursor" & _
			" are visible. I will now select the data the TextCursor has selected.")

	; Select the data the TextCursor has selected
	_LOWriter_ParObjSelect($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to select the TextCursor Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
