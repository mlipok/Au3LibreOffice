#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oCellRange
	Local $avColor[0], $avColor2[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell B5
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set Cell B5's Background color to yellow.
	_LOCalc_CellBackColor($oCell, $LOC_COLOR_YELLOW)
	If @error Then _ERROR($oDoc, "Failed to set Cell Background color. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell range A1 to A6
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Set Cell A1-A6's Background color to Blue.
	_LOCalc_CellBackColor($oCellRange, $LOC_COLOR_BLUE)
	If @error Then _ERROR($oDoc, "Failed to set Cell Range Background color. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell B5's current background color setting, Return will be an array in order of function parameters.
	$avColor = _LOCalc_CellBackColor($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Background color. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell Range A1-A6's current background color setting, Return will be an array in order of function parameters.
	$avColor2 = _LOCalc_CellBackColor($oCellRange)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Background color. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Cell B5's Background color is, in Long integer format: " & $avColor[0] & @CRLF & _
			"Is Cell B5's Background color transparent? True/False: " & $avColor[1] & @CRLF & _
			"Cell Range A1-A6's Background color is, in Long integer format: " & $avColor2[0] & @CRLF & _
			"Is Cell Range A1-A6's Background color transparent? True/False: " & $avColor2[1])

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
