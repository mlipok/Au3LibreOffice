#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellStyle, $oCell
	Local $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell A1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Insert some text into Cell A1
	_LOCalc_CellString($oCell, "Cell A1")
	If @error Then _ERROR($oDoc, "Failed to set Cell Text. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell B2
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Insert some text into Cell B2
	_LOCalc_CellString($oCell, "Some Text")
	If @error Then _ERROR($oDoc, "Failed to set Cell Text. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Object for "Default" Cell Style.
	$oCellStyle = _LOCalc_CellStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Style Object. Error:" & @error & " Extended:" & @extended)

	; Set the Cell Style's Text properties to, Wrap Text Automatically = True, Auto Hyphenate = True, Skip Shrink to fit, Text direction = $LOC_TXT_DIR_RL .
	_LOCalc_CellStyleTextProperties($oCellStyle, True, True, Null, $LOC_TXT_DIR_RL)
	If @error Then _ERROR($oDoc, "Failed to set the Cell Style's settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOCalc_CellStyleTextProperties($oCellStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Cell Style's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Cell Style's current Text properties are as follows: " & @CRLF & _
			"Is Text automatically wrapped? True/False: " & $avSettings[0] & @CRLF & _
			"Is Text automatically hyphenated? True/False: " & $avSettings[1] & @CRLF & _
			"Is Text font size automatically shrunk to fit the cell width? True/False: " & $avSettings[2] & @CRLF & _
			"The Cell Style Text direction is (See UDF Constants): " & $avSettings[3])

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
