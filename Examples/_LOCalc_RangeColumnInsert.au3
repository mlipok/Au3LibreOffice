#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oCellRange
	Local $iCount = 0

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the Cell Range of A1 to F3 with numbers, one cell at a time. (Remember Columns and Rows are 0 based.)
	For $i = 0 To 5
		For $j = 0 To 2
			; Retrieve the Cell Object
			$oCell = _LOCalc_RangeGetCellByPosition($oSheet, $i, $j)
			If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			; Set the Cell to a Number
			_LOCalc_CellValue($oCell, $iCount)
			If @error Then _ERROR($oDoc, "Failed to set Cell Value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			$iCount += 1
		Next
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now Insert 2 columns Starting at B.")

	; Insert new Columns at B, B is column 1 because L.O. Columns are 0 based.
	_LOCalc_RangeColumnInsert($oSheet, 1, 2)
	If @error Then _ERROR($oDoc, "Failed to insert columns. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now Retrieve Cell Range D3 to F7, and insert a new column in the second over Column position.")

	; Retrieve Cell Range D3 to F7.
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "D3", "F7")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Column at second over column position
	_LOCalc_RangeColumnInsert($oCellRange, 1)
	If @error Then _ERROR($oDoc, "Failed to insert columns. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
