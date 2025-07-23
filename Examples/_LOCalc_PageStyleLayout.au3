#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOCalc_PageStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page layout to, $LOC_PAGE_LAYOUT_MIRRORED, Numbering format to $LOC_NUM_STYLE_CHARS_UPPER_LETTER_N, Align Table Horizontally = True,
	; Align Table Vertical = True
	_LOCalc_PageStyleLayout($oPageStyle, $LOC_PAGE_LAYOUT_MIRRORED, $LOC_NUM_STYLE_CHARS_UPPER_LETTER_N, True, True)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avPageStyleSettings = _LOCalc_PageStyleLayout($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Layout settings are as follows: " & @CRLF & _
			"The current Page Layout is, (see UDF constants): " & $avPageStyleSettings[0] & @CRLF & _
			"The Numbering format used is, (See UDF constants): " & $avPageStyleSettings[1] & @CRLF & _
			"Will the Table be Aligned to the center Horizontally? True/False: " & $avPageStyleSettings[2] & @CRLF & _
			"Will the Table be Aligned to the center Vertically? True/False: " & $avPageStyleSettings[3] & @CRLF & _
			"The paper tray to use, when printing this document is: " & $avPageStyleSettings[4])

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
