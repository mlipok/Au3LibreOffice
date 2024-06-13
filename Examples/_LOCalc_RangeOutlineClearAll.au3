#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oCell
	Local $aavData[1]
	Local $avRowData[8]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Fill my arrays with the desired Number Values I want in Column A.
	$avRowData[0] = 1 ; A1
	$avRowData[1] = 2 ; B1
	$avRowData[2] = 3 ; C1
	$avRowData[3] = 0 ; D1
	$avRowData[4] = 10 ; E1
	$avRowData[5] = 15 ; F1
	$avRowData[6] = 25 ; G1
	$avRowData[7] = 0 ; H1
	$aavData[0] = $avRowData

	; Retrieve Cell range A1 to F1
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "F1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Fill the range with Data
	_LOCalc_RangeNumbers($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell  D1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "D1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set formula for Cell D1
	_LOCalc_CellFormula($oCell, "=SUM(A1:C1)")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell  H1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "H1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set formula for Cell H1
	_LOCalc_CellFormula($oCell, "=SUM(E1:G1)")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Set up AutoOutline for the Range.
	_LOCalc_RangeAutoOutline($oCellRange)
	If @error Then _ERROR($oDoc, "Failed to set up Auto Outline for the Range. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell Range B2 to B5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B2", "B5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Group the Rows of B2 to 5
	_LOCalc_RangeGroup($oCellRange, $LOC_GROUP_ORIENT_ROWS, True)
	If @error Then _ERROR($oDoc, "Failed to group Cell range. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to clear all Outline Groups.")

	; Clear all Outline Groups for the Sheet.
	_LOCalc_RangeOutlineClearAll($oSheet)
	If @error Then _ERROR($oDoc, "Failed to Clear Cell Groups. Error:" & @error & " Extended:" & @extended)

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
