#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField, $oFieldAnchor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Count Field at the View Cursor. Set count type to $LOW_FIELD_COUNT_TYPE_WORDS, Overwrite to False, and Number format to
	; $LOW_NUM_STYLE_CHARS_LOWER_LETTER
	$oField = _LOWriter_FieldStatCountInsert($oDoc, $oViewCursor, $LOW_FIELD_COUNT_TYPE_WORDS, False, $LOW_NUM_STYLE_ARABIC)
	If @error Then _ERROR($oDoc, "Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Field's anchor (Text Cursor)
	$oFieldAnchor = _LOWriter_FieldGetAnchor($oField)
	If @error Then _ERROR($oDoc, "Failed to retrieve a Field Anchor. Error:" & @error & " Extended:" & @extended)

	; Move the Anchor (Text Cursor)
	_LOWriter_CursorMove($oFieldAnchor, $LOW_TEXTCUR_GO_RIGHT, 1, False)
	If @error Then _ERROR($oDoc, "Failed to move a cursor. Error:" & @error & " Extended:" & @extended)

	; Insert Some text.
	_LOWriter_DocInsertString($oDoc, $oFieldAnchor, " Some new text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

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
