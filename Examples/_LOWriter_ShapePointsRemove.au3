#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $atArray[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Polygon Shape into the document, 3000 Wide by 2000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_LINE_POLYGON, 3000, 2000)
	If @error Then _ERROR("Failed to create a Shape. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Shape's current Array of Points.
	$atArray = _LOWriter_ShapePoints($oShape)
	If @error Then _ERROR("Failed to retrieve Array of Shape points. Error:" & @error & " Extended:" & @extended)

	; Delete the third Position Point
	_LOWriter_ShapePointsRemove($atArray, 2)
	If @error Then _ERROR("Failed to modify Shape point. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to insert the modified Points into the shape.")

	; Re-insert the modified Points
	_LOWriter_ShapePoints($oShape, $atArray)
	If @error Then _ERROR("Failed to modify Array of Shape points. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
