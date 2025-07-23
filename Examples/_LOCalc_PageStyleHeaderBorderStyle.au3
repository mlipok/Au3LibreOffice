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

	; Turn Header on.
	_LOCalc_PageStyleHeader($oPageStyle, True)
	If @error Then _ERROR($oDoc, "Failed to turn Page Style headers on. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Header Border Width (all four sides) to $LOC_BORDERWIDTH_MEDIUM
	_LOCalc_PageStyleHeaderBorderWidth($oPageStyle, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM, $LOC_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Header Border Style settings to: Top = $LOC_BORDERSTYLE_DASH_DOT_DOT, Bottom = $LOC_BORDERSTYLE_THICKTHIN_MEDIUMGAP
	; Left = $LOC_BORDERSTYLE_DOUBLE, Right = $LOC_BORDERSTYLE_DASHED
	_LOCalc_PageStyleHeaderBorderStyle($oPageStyle, $LOC_BORDERSTYLE_DASH_DOT_DOT, $LOC_BORDERSTYLE_THICKTHIN_MEDIUMGAP, $LOC_BORDERSTYLE_DOUBLE, $LOC_BORDERSTYLE_DASHED)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOCalc_PageStyleHeaderBorderStyle($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Header Border Style settings are as follows: " & @CRLF & _
			"The Top Border Style is, (see UDF constants): " & $avPageStyleSettings[0] & @CRLF & _
			"The Bottom Border Style is, (see UDF constants): " & $avPageStyleSettings[1] & @CRLF & _
			"The Left Border Style is, (see UDF constants): " & $avPageStyleSettings[2] & @CRLF & _
			"The Right Border Style is, (see UDF constants): " & $avPageStyleSettings[3])

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
