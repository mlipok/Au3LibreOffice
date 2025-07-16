#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oParStyle
	Local $iMicrometers, $iTabStop
	Local $avTabStop

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the "Default Paragraph Style" object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Default Paragraph Style")
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Tab Stop for the demonstration.
	$iTabStop = _LOWriter_ParStyleTabStopCreate($oParStyle, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.5)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the TabStop from 1/4" to 1/2" Tab Stop position, Set the fill character to Asc(~) the Tilde key ASCII Value 126.
	; Set alignment To  $LOW_TAB_ALIGN_DECIMAL, and the decimal character to ASC(.) a period, ASCII value 46.
	; Since I am modifying the TabStop position, @Extended will be my new Tab Stop position.
	_LOWriter_ParStyleTabStopMod($oParStyle, $iTabStop, $iMicrometers, Asc("~"), $LOW_TAB_ALIGN_DECIMAL, Asc("."))
	$iTabStop = @extended
	If @error Then _ERROR($oDoc, "Failed to modify Paragraph Tab stop settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avTabStop = _LOWriter_ParStyleTabStopMod($oParStyle, $iTabStop)
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph Tab stop settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Tab stop, having the position of " & $iTabStop & " has the following settings: " & @CRLF & _
			"The Current position is, in micrometers: " & $avTabStop[0] & @CRLF & _
			"The Current Fill Character is, in ASC value: " & $avTabStop[1] & " and looks like: " & Chr($avTabStop[1]) & @CRLF & _
			"The Current Alignment setting is, (See UDF constants): " & $avTabStop[2] & @CRLF & _
			"The Current Decimal Character is, in ASC value: " & $avTabStop[3] & " and looks like: " & Chr($avTabStop[3]))

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
