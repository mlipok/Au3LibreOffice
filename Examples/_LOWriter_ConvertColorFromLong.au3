#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $sHex
	Local $aiRGB, $aiCMYK, $aiHSB

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 5 rows, 3 columns.
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert the Table into the document at the View Cursor's location.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR($oDoc, "Failed to insert Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the table Background color to and set Transparent to False.
	_LOWriter_TableColor($oTable, $LOW_COLOR_MAGENTA, False)
	If @error Then _ERROR($oDoc, "Failed to set Text Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I am going to demonstrate how to convert the Long color format integer value, $LOW_COLOR_MAGENTA (12517441), into R(ed), G(reen), " & _
			"B(lue) values, a Hexadecimal value, C(yan), M(agenta), Y(ellow), and K(ey) values, and H(ue), S(aturation) B(rightness) values.")

	; Convert to RGB From Long Color format, the RGB values are returned as an array in their order.
	$aiRGB = _LOWriter_ConvertColorFromLong(Null, $LOW_COLOR_MAGENTA)
	If @error Then _ERROR($oDoc, "Failed to convert to RGB color value from Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert to Hex From Long color format, Hex is returned as a string.
	$sHex = _LOWriter_ConvertColorFromLong($LOW_COLOR_MAGENTA)
	If @error Then _ERROR($oDoc, "Failed to convert to HEX color value from Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert to CMYK From Long Color format, the CMYK values are returned as an array in their order.
	$aiCMYK = _LOWriter_ConvertColorFromLong(Null, Null, Null, $LOW_COLOR_MAGENTA)
	If @error Then _ERROR($oDoc, "Failed to convert to CMYK color value from Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert to HSB From Long Color format, the HSB values are returned as an array in their order.
	$aiHSB = _LOWriter_ConvertColorFromLong(Null, Null, $LOW_COLOR_MAGENTA)
	If @error Then _ERROR($oDoc, "Failed to convert to HSB color value from Long color format integer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The conversion results are as follows: " & @CRLF & _
			"Long->RGB = " & "R, " & $aiRGB[0] & "; G, " & $aiRGB[1] & "; B, " & $aiRGB[2] & @CRLF & " Should be: R, 191; G, 0; B, 65" & @CRLF & @CRLF & _
			"Long->Hex = " & $sHex & @CRLF & " Should be bf0041" & @CRLF & @CRLF & _
			"Long->CMYK = " & "C, " & $aiCMYK[0] & "; M " & $aiCMYK[1] & "; Y " & $aiCMYK[2] & "; K " & $aiCMYK[3] & @CRLF & " Should be: C, 0; M, 100; Y, 66; K, 25." & @CRLF & @CRLF & _
			"Long->HSB = " & "H, " & $aiHSB[0] & "; S " & $aiHSB[1] & "; B " & $aiHSB[2] & @CRLF & " Should be: H, 340; S, 100; B, 75")

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
