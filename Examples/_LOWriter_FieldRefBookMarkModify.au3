#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Bookmark at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Bookmark at the ViewCursor, named "New Bookmark".
	_LOWriter_DocBookmarkInsert($oDoc, $oViewCursor, False, "New Bookmark")
	If @error Then _ERROR($oDoc, "Failed to insert a Bookmark. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Bookmark Reference Field. "New Bookmark" Bookmark, do not overwrite any content selected by the cursor, and refer using
	; $LOW_FIELD_REF_USING_REF_TEXT
	$oField = _LOWriter_FieldRefBookMarkInsert($oDoc, $oViewCursor, "New Bookmark", False, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED)

	; Insert some more text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "I have inserted a Bookmark at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Book Mark at the ViewCursor, named "Second Bookmark".
	_LOWriter_DocBookmarkInsert($oDoc, $oViewCursor, False, "Second Bookmark")
	If @error Then _ERROR($oDoc, "Failed to insert a Book Mark. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to modify the Reference Field.")

	; Modify the Bookmark Reference Field settings. Set the Referenced Bookmark to "Second Bookmark", and Refer using $LOW_FIELD_REF_USING_ABOVE_BELOW
	_LOWriter_FieldRefBookMarkModify($oDoc, $oField, "Second Bookmark", $LOW_FIELD_REF_USING_ABOVE_BELOW)
	If @error Then _ERROR($oDoc, "Failed to modify field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Field settings. Return will be an Array with elements in the order of function parameters.
	$avSettings = _LOWriter_FieldRefBookMarkModify($oDoc, $oField)
	If @error Then _ERROR($oDoc, "Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Field settings are: " & @CRLF & _
			"The Bookmark name being referenced is: " & $avSettings[0] & @CRLF & _
			"The Bookmark is being referenced using this format, (see UDF Constants): " & $avSettings[1])

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
