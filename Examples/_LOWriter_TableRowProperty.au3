#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $iRows, $iColumns, $iMicrometers
	Local $avRowProps

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

	; Retrieve how many Rows the Table currently contains.
	$iRows = _LOWriter_TableRowGetCount($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Row count. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve how many Columns the Table currently contains.
	$iColumns = _LOWriter_TableColumnGetCount($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Column count. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $iRow = 0 To $iRows - 1
		For $iColumn = 0 To $iColumns - 1
			; Retrieve each cell by position in the Table.
			$oCell = _LOWriter_TableGetCellObjByPosition($oTable, $iColumn, $iRow)
			If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Cell by position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			; Set Cell text String to Cell's position.
			_LOWriter_CellString($oCell, "Column " & $iColumn & @CR & " Row " & $iRow)
			If @error Then _ERROR($oDoc, "Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
		Next
	Next

	; Retrieve the third down (Row 2) settings.
	$avRowProps = _LOWriter_TableRowProperty($oTable, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve row settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Array elements will be in order of function's parameters.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The settings for the Third row down, (Row 2), are: " & @CRLF & _
			"Height: " & $avRowProps[0] & @CRLF & _
			"Is Auto Height? True/False: " & $avRowProps[1] & @CRLF & _
			"Is the Row allowed to split between pages? True/False: " & $avRowProps[2])

	; Convert 1.5 inches to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(1.5)
	If @error Then _ERROR($oDoc, "Failed to convert inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Row properties to: 1.5" height, turn off auto height, and disallow splitting the row between pages.
	_LOWriter_TableRowProperty($oTable, 2, $iMicrometers, False, False)
	If @error Then _ERROR($oDoc, "Failed to set row properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the third down (Row 2) settings again.
	$avRowProps = _LOWriter_TableRowProperty($oTable, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve row settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Array elements will be in order of function's parameters.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The new settings for the Third row down, (Row 2), are: " & @CRLF & _
			"Height: " & $avRowProps[0] & @CRLF & _
			"Is Auto Height? True/False: " & $avRowProps[1] & @CRLF & _
			"Is the Row allowed to split between pages? True/False: " & $avRowProps[2])

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
