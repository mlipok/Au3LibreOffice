#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oRow
	Local $aoRanges[0]
	Local $iResults

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Row Object for Row 2 (numbers Row 1 because L.O. Rows are internally 0 based.)
	$oRow = _LOCalc_RangeRowGetObjByPosition($oSheet, 1)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Object for Row 2 (1). Error:" & @error & " Extended:" & @extended)

	; Set Row 2's Visibility to False (Invisible).
	_LOCalc_RangeRowVisible($oRow, False)
	If @error Then _ERROR($oDoc, "Failed to row's Visibility. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell Range A1-C3
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Query the Cell Range for Visible cells.
	$aoRanges = _LOCalc_RangeQueryVisible($oCellRange)
	If @error Then _ERROR($oDoc, "Failed to Query Cell for Visible Cells. Error:" & @error & " Extended:" & @extended)
	$iResults = @extended

	MsgBox($MB_OK, "", "I will now highlight in yellow the cell ranges that are Visible in the Cell Range, and then set Row 2 to visible again. Take note that Row 2 is not highlighted.")

	; Cycle through the results and set the background color to yellow for each Cell range found
	For $i = 0 To $iResults - 1
		_LOCalc_CellBackColor($aoRanges[$i], $LOC_COLOR_YELLOW)
		If @error Then _ERROR($oDoc, "Failed to set Cell Range Background color. Error:" & @error & " Extended:" & @extended)
	Next

	; Set Row 2's Visibility to True (Visible).
	_LOCalc_RangeRowVisible($oRow, True)
	If @error Then _ERROR($oDoc, "Failed to row's Visibility. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
