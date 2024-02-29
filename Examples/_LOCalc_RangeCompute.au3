#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $aavData[5]
	Local $avRowData[1]
	Local $nResult

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Fill my arrays with the desired Number Values I want in Column A.
	$avRowData[0] = 1 ; A1
	$aavData[0] = $avRowData

	$avRowData[0] = 2.25 ; A2
	$aavData[1] = $avRowData

	$avRowData[0] = 523.89 ; A3
	$aavData[2] = $avRowData

	$avRowData[0] = 18 ; A4
	$aavData[3] = $avRowData

	$avRowData[0] = 537     ; A5
	$aavData[4] = $avRowData

	; Retrieve Cell range A1 to A5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Fill the range with Data
	_LOCalc_RangeNumbers($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended)

	; Perform a SUM Compute Function on the Range.
	$nResult = _LOCalc_RangeCompute($oCellRange, $LOC_COMPUTE_SUM)
	If @error Then _ERROR($oDoc, "Failed to Compute Cell Range. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The SUM of the Cell Range is: " & $nResult)

	; Perform an AVERAGE Compute Function on the Range.
	$nResult = _LOCalc_RangeCompute($oCellRange, $LOC_COMPUTE_AVERAGE)
	If @error Then _ERROR($oDoc, "Failed to Compute Cell Range. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The AVERAGE of the Cell Range is: " & $nResult)

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
