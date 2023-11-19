#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $atArray[0], $atArray2[0]
	Local $iNewX, $iNewY

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a FreeForm Line Shape into the document, 3000 Wide by 2000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_LINE_FREEFORM_LINE, 3000, 2000)
	If @error Then _ERROR("Failed to create a Shape. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Shape's current Array of Points.
	$atArray = _LOWriter_ShapePoints($oShape)
	If @error Then _ERROR("Failed to retrieve Array of Shape points. Error:" & @error & " Extended:" & @extended)

	; I will add a point after the the second Position Point in the line, so I am going to retrieve the second points current position, and add to its X and Y values to
	; determine my third point's new values.

	; Retrieve the second point's X and Y settings.
	$atArray2 = _LOWriter_ShapePointsModify($atArray, 1)
	If @error Then _ERROR("Failed to retrieve Shape point settings. Error:" & @error & " Extended:" & @extended)

	; Minus 200 Micrometers from the X coordinate
	$iNewx = $atArray2[0] - 200

	; Add 300 Micrometers to the Y coordinate
	$iNewY = $atArray2[1] + 300

	; Add the new Point using the new X and Y coordinates. The new point will be added before the called array element so I add 1 to the desired placement value.
	_LOWriter_ShapePointsAdd($atArray, 2, $iNewx, $iNewY)
	If @error Then _ERROR("Failed to modify Shape point. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to insert the new Point into the shape.")

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
