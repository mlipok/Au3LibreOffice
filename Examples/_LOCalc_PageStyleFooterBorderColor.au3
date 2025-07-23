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

	; Turn Footer on.
	_LOCalc_PageStyleFooter($oPageStyle, True)
	If @error Then _ERROR($oDoc, "Failed to turn Page Style footers on. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Footer Border Width (all four sides) to $LOC_BORDERWIDTH_MEDIUM
	_LOCalc_PageStyleFooterBorderWidth($oPageStyle, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Footer Border Color settings to: Top, $LOC_COLOR_ORANGE, Bottom $LOC_COLOR_BLUE, Left, $LOC_COLOR_LGRAY, Right $LOC_COLOR_BLACK
	_LOCalc_PageStyleFooterBorderColor($oPageStyle, $LOC_COLOR_ORANGE, $LOC_COLOR_BLUE, $LOC_COLOR_LGRAY, $LOC_COLOR_BLACK)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOCalc_PageStyleFooterBorderColor($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Footer Border Color settings are as follows: " & @CRLF & _
			"The Top Border Color is, in Long Color Format: " & $avPageStyleSettings[0] & @CRLF & _
			"The Bottom Border Color is, in Long Color Format: " & $avPageStyleSettings[1] & @CRLF & _
			"The Left Border Color is, in Long Color Format: " & $avPageStyleSettings[2] & @CRLF & _
			"The Right Border Color is, in Long Color Format: " & $avPageStyleSettings[3])

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
