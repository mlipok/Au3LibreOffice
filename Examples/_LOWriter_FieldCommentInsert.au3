#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $tDateStruct

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line. ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Create a Date Structure, Year = 1844, Month = 10, Day = 22, Hour = 8, minutes = 14, Seconds = 0 , Nanoseconds = 0, UTC= True.
	$tDateStruct = _LOWriter_DateStructCreate(1844, 10, 22, 8, 14, 0, 0, True)
	If @error Then _ERROR($oDoc, "Failed to create a Date structure. Error:" & @error & " Extended:" & @extended)

	; Insert a Comment Field at the View Cursor. Set content to "This is a note", Author to "Daniel", Date to my previous Date Structure I created.
	; Initials to "D.", Author to "A Name", Resolved = False
	_LOWriter_FieldCommentInsert($oDoc, $oViewCursor, False, "This is a note", "Daniel", $tDateStruct, "D.", "A Name", False)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended)

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
