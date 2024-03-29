#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame, $oTextCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Frame into the document at the ViewCursor position, and 5000x5000 Micrometers wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 5000, 5000)
	If @error Then _ERROR($oDoc, "Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	; Create a Text Cursor in the frame.
	$oTextCursor = _LOWriter_FrameCreateTextCursor($oFrame)
	If @error Then _ERROR($oDoc, "Failed to create a text Cursor Object for the Frame. Error:" & @error & " Extended:" & @extended)

	; Insert some text
	_LOWriter_DocInsertString($oDoc, $oTextCursor, "First Line of text." & @CR & _
			"Second line of text.")
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
