#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oCell
	Local $iCount = 0

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Fill the Cell Range of A1 to D5 with numbers, one cell at a time. (Remember Columns and Rows are 0 based.)
	For $i = 0 To 4
		For $j = 0 To 4
			; Retrieve the Cell Object
			$oCell = _LOCalc_SheetGetCellByPosition($oSheet, $i, $j)
			If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

			; Set the Cell to a Number
			_LOCalc_CellValue($oCell, $iCount)
			If @error Then _ERROR($oDoc, "Failed to set Cell Value. Error:" & @error & " Extended:" & @extended)

			$iCount += 1

		Next

	Next

	; Retrieve the Cell Range B2 to C3
	$oCellRange = _LOCalc_SheetGetCellByName($oSheet, "B2", "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now delete the range B2 to C3, shifting the cells below up.")

	; Delete the Cell Range
	_LOCalc_CellRangeDelete($oSheet, $oCellRange, $LOC_CELL_DELETE_MODE_UP)
	If @error Then _ERROR($oDoc, "Failed to delete Cell Range Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now delete the range B2 to C3 again, shifting the cells right of the range, to the left.")

	; Delete the Cell Range
	_LOCalc_CellRangeDelete($oSheet, $oCellRange, $LOC_CELL_DELETE_MODE_LEFT)
	If @error Then _ERROR($oDoc, "Failed to delete Cell Range Object. Error:" & @error & " Extended:" & @extended)

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
