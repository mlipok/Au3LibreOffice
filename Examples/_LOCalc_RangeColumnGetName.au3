#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oColumn, $oCellRange
	Local $sName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the third over column, which is column C. Remember Columns are 0 based.
	$oColumn = _LOCalc_RangeColumnGetObjByPosition($oSheet, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Column's Name.
	$sName = _LOCalc_RangeColumnGetName($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column's name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have retrieved Column #2, (The third column in the Sheet), the Column's name is: " & $sName & @CRLF & @CRLF & _
			"I will now retrieve Cell Range C1 to F5 and retrieve the third Column in the Cell Range.")

	; Retrieve Cell Range C1 to F5.
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "C1", "F5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the third over column.
	$oColumn = _LOCalc_RangeColumnGetObjByPosition($oCellRange, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Column's Name.
	$sName = _LOCalc_RangeColumnGetName($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column's name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have retrieved Column #2, (The third column in the Cell Range), the Column's name is: " & $sName)

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
