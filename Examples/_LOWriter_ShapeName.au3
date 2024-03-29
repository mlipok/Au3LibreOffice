#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $sName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Rectangle Shape into the document, 3000 Wide by 6000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Shape name
	$sName = _LOWriter_ShapeName($oDoc, $oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "Name", "The Shape's current name is: " & $sName)

	; Change the Shape's name to "AutoIt Test"
	$sName = _LOWriter_ShapeName($oDoc, $oShape, "AutoIt Test")
	If @error Then _ERROR($oDoc, "Failed to modify Shape name. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Shape name
	$sName = _LOWriter_ShapeName($oDoc, $oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "Name", "The Shape's new name is: " & $sName)

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
