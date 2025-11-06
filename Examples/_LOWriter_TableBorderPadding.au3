#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable
	Local $i100thMM
	Local $aiPadding

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create the Table, 5 rows, 3 columns
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR($oDoc, "Failed to insert Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the outer table borders to medium thickness.
	_LOWriter_TableBorderWidth($oTable, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR($oDoc, "Failed to set Table Border width settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2 Inch to Hundredths of a Millimeter (100th MM).
	$i100thMM = _LO_UnitConvert(0.5, $LO_CONVERT_UNIT_INCH_100THMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (100th MM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Table padding to 1/2"
	_LOWriter_TableBorderPadding($oTable, $i100thMM, $i100thMM, $i100thMM, $i100thMM)
	If @error Then _ERROR($oDoc, "Failed to set Table Border Padding settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Table Border Padding settings. Return will be an Array, with values in order of function parameters.
	$aiPadding = _LOWriter_TableBorderPadding($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table Border Padding settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Table Border padding values are as follows: " & @CRLF & _
			"Top = " & $aiPadding[0] & " Hundredths of a Millimeter (100th MM)" & @CRLF & _
			"Bottom = " & $aiPadding[1] & " Hundredths of a Millimeter (100th MM)" & @CRLF & _
			"Left = " & $aiPadding[2] & " Hundredths of a Millimeter (100th MM)" & @CRLF & _
			"Right = " & $aiPadding[3] & " Hundredths of a Millimeter (100th MM)")

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
