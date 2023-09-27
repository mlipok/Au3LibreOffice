#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame, $oTextCursor, $oFrameNew

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Frame into the document at the Viewcursor position, and 3000x3000 Micrometers wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	; Create a Text Cursor in the frame.
	$oTextCursor = _LOWriter_FrameCreateTextCursor($oFrame)
	If (@error > 0) Then _ERROR("Failed to create a text Cursor Object for the Frame. Error:" & @error & " Extended:" & @extended)

	; Move the View Curosr into the frame.
	_LOWriter_CursorGoToRange($oViewCursor, $oTextCursor)
	If (@error > 0) Then _ERROR("Failed to move a Cursor Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Frame Object by the ViewCursor (which is located in the Frame.)
	$oFrameNew = _LOWriter_FrameGetObjByCursor($oDoc, $oViewCursor)
	If (@error > 0) Then _ERROR("Failed to Retrieve a Frame Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to delete the Frame, using the newly retrieved Object.")

	; Delete the frame using the new Frame Object.
	_LOWriter_FrameDelete($oDoc, $oFrameNew)
	If (@error > 0) Then _ERROR("Failed to delete the Frame. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
