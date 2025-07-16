#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $iCount
	Local $asFonts

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array list of font names
	$asFonts = _LOWriter_FontsGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Array of font names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iCount = @extended

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were " & $iCount & " fonts found. I will now insert the results into a table. The Table will have four columns, " & @CRLF & _
			"-the first column contains the font name, " & @CRLF & _
			"-the second column contains the style name, " & @CRLF & _
			"-the third column contains the Font weight (Bold) value, (see constants)," & @CRLF & _
			"-the fourth column contains the font slant (Italic), (See constants).")

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table
	$oTable = _LOWriter_TableCreate($oDoc, $iCount, 4)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR($oDoc, "Failed to insert Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $iRow = 0 To $iCount - 1
		; Retrieve each cell by position in the Table.
		$oCell = _LOWriter_TableGetCellObjByPosition($oTable, 0, $iRow)
		If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Cell by position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set Cell text String to Font's Name.
		_LOWriter_CellString($oCell, $asFonts[$iRow][0])
		If @error Then _ERROR($oDoc, "Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve each cell by position in the Table.
		$oCell = _LOWriter_TableGetCellObjByPosition($oTable, 1, $iRow)
		If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Cell by position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set Cell text String to Font's Style Name.
		_LOWriter_CellString($oCell, $asFonts[$iRow][1])
		If @error Then _ERROR($oDoc, "Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve each cell by position in the Table.
		$oCell = _LOWriter_TableGetCellObjByPosition($oTable, 2, $iRow)
		If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Cell by position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set Cell text String to Font's Weight.
		_LOWriter_CellValue($oCell, $asFonts[$iRow][2])
		If @error Then _ERROR($oDoc, "Failed to set Text Table Cell Value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve each cell by position in the Table.
		$oCell = _LOWriter_TableGetCellObjByPosition($oTable, 3, $iRow)
		If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Cell by position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set Cell text String to Font's Italic setting.
		_LOWriter_CellValue($oCell, $asFonts[$iRow][3])
		If @error Then _ERROR($oDoc, "Failed to set Text Table Cell Value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

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
