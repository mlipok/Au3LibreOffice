#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oColumn
	Local $bVisible

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Column D's Object.
	$oColumn = _LOCalc_SheetColumnGetObjByName($oSheet, "D")
	If @error Then _ERROR($oDoc, "Failed to retrieve the Column Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Column D's current visibility setting.
	$bVisible = _LOCalc_SheetColumnVisible($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column's current visibility setting. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Column D currently visible? True/False: " & $bVisible)

	; Set Column D to invisible.
	_LOCalc_SheetColumnVisible($oColumn, False)
	If @error Then _ERROR($oDoc, "Failed to set Column's visibility setting. Error:" & @error & " Extended:" & @extended)

	; Retrieve Column D's current visibility setting again.
	$bVisible = _LOCalc_SheetColumnVisible($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column's current visibility setting. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Now is Column D visible? True/False: " & $bVisible)

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
