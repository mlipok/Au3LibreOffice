#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Turn Header on.
	_LOWriter_PageStyleHeader($oPageStyle, True)
	If @error Then _ERROR($oDoc, "Failed to turn Page Style headers on. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Page style Header Gradient settings to: Preset Gradient name = $LOW_GRAD_NAME_TEAL_TO_BLUE
	_LOWriter_PageStyleHeaderAreaGradient($oDoc, $oPageStyle, $LOW_GRAD_NAME_TEAL_TO_BLUE)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an integer.
	$avPageStyleSettings = _LOWriter_PageStyleHeaderAreaGradient($oDoc, $oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Header Gradient settings are as follows: " & @CRLF & _
			"The Gradient name is: " & $avPageStyleSettings[0] & @CRLF & _
			"The type of Gradient is, (see UDF constants): " & $avPageStyleSettings[1] & @CRLF & _
			"The number of steps to increment color is: " & $avPageStyleSettings[2] & @CRLF & _
			"The horizontal offset percentage for the gradient is: " & $avPageStyleSettings[3] & @CRLF & _
			"The vertical offset percentage for the gradient is: " & $avPageStyleSettings[4] & @CRLF & _
			"The rotation angle for the gradient is, in degrees: " & $avPageStyleSettings[5] & @CRLF & _
			"The percentage of area not covered by the ending color is: " & $avPageStyleSettings[6] & @CRLF & _
			"The starting color is, in Long Color format: " & $avPageStyleSettings[7] & @CRLF & _
			"The ending color is, in Long Color format: " & $avPageStyleSettings[8] & @CRLF & _
			"The starting color intensity percentage is: " & $avPageStyleSettings[9] & @CRLF & _
			"The ending color intensity percentage is: " & $avPageStyleSettings[10])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
