#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iMicrometers
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	; Convert 1/8" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.125)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Set Page style Shadow settings to: Width = 1/8", Color = $LOW_COLOR_RED, Transparent = False, Location = $LOW_SHADOW_TOP_LEFT
	_LOWriter_PageStyleShadow($oPageStyle, $iMicrometers, $LOW_COLOR_RED, False, $LOW_SHADOW_TOP_LEFT)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleShadow($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Shadow settings are as follows: " & @CRLF & _
			"The shadow width is, is Micrometers: " & $avPageStyleSettings[0] & @CRLF & _
			"The Shadow color is, in Long Color format: " & $avPageStyleSettings[1] & @CRLF & _
			"Is the Color transparent? True/False: " & $avPageStyleSettings[2] & @CRLF & _
			"The Shadow location is, (see UDF Constants): " & $avPageStyleSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
