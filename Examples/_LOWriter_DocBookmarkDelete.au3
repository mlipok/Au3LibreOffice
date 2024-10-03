#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oBookmark

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
	$oBookmark = _LOWriter_DocBookmarkInsert($oDoc, $oViewCursor, False, "New Bookmark")
	If @error Then _ERROR($oDoc, "Failed to insert a Bookmark. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK, "", "Does the document contain a Bookmark named ""New Bookmark""? True/False: " & _LOWriter_DocBookmarksHasName($oDoc, "New Bookmark") & _
			@CRLF & @CRLF & "Press Ok to delete the Bookmark.")

	; Delete the Bookmark.
	_LOWriter_DocBookmarkDelete($oDoc, $oBookmark)
	If @error Then _ERROR($oDoc, "Failed to delete a Bookmark. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK, "", "Now does the document contain a Bookmark named ""New Bookmark""? True/False: " & _LOWriter_DocBookmarksHasName($oDoc, "New Bookmark"))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
