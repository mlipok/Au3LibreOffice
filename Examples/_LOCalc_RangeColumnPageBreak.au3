#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oColumn
	Local $abSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Column E's Object.
	$oColumn = _LOCalc_RangeColumnGetObjByName($oSheet, "E")
	If @error Then _ERROR($oDoc, "Failed to retrieve the Column Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the column page break
	_LOCalc_RangeColumnPageBreak($oColumn, True)
	If @error Then _ERROR($oDoc, "Failed to set Column Page Break Settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have inserted a manual Page Break at Column E.")

	; Retrieve the Page Break Settings for Column E. Return will be an array with setting values in order of Function parameters.
	$abSettings = _LOCalc_RangeColumnPageBreak($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Page Break Settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Column E's current Page Break settings are:" & @CRLF & _
			"Is there a Manual Page Break at Column E? True/False: " & $abSettings[0] & @CRLF & _
			"Is Column E the start of a new Page? True/False: " & $abSettings[1])

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
