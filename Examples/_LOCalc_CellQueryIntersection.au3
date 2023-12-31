#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oCellRange2
	Local $aoRanges[0]
	Local $iResults

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell Range A1-C6
	$oCellRange = _LOCalc_SheetGetCellByName($oSheet, "A1", "C6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Set the background Color of Range A1-C6 to Blue
	_LOCalc_CellBackColor($oCellRange, $LOC_COLOR_BLUE)
	If @error Then _ERROR($oDoc, "Failed to set Cell Range Background color. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell Range B3-C3
	$oCellRange2 = _LOCalc_SheetGetCellByName($oSheet, "B3", "E5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Set the background Color of Range B3-E5 to Red
	_LOCalc_CellBackColor($oCellRange2, $LOC_COLOR_RED)
	If @error Then _ERROR($oDoc, "Failed to set Cell Range Background color. Error:" & @error & " Extended:" & @extended)

	; Query Cell Range 1 and Cell Range 2 for intersecting cells .
	$aoRanges = _LOCalc_CellQueryIntersection($oCellRange, $oCellRange2)
	If @error Then _ERROR($oDoc, "Failed to Query Cells that Intersect in the ranges. Error:" & @error & " Extended:" & @extended)
	$iResults = @extended

	MsgBox($MB_OK, "", "I will now highlight in yellow the cell ranges that are intersecting between Cell Range 1 and Cell Range 2.")

	; Cycle through the results and set the background color to yellow for each Cell range found
	For $i = 0 To $iResults - 1
		_LOCalc_CellBackColor($aoRanges[$i], $LOC_COLOR_YELLOW)
		If @error Then _ERROR($oDoc, "Failed to set Cell Range Background color. Error:" & @error & " Extended:" & @extended)
	Next

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
