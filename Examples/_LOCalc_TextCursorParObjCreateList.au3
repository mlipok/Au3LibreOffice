#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oTextCursor, $oRow
	Local $aoPar[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve active Sheet's Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell A1's Object
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Create a Text Cursor in the Cell
	$oTextCursor = _LOCalc_CellCreateTextCursor($oCell)
	If @error Then _ERROR($oDoc, "Failed to create Text Cursor in Cell. Error:" & @error & " Extended:" & @extended)

	; Insert some words
	_LOCalc_TextCursorInsertString($oTextCursor, "Hi!" & @CRLF & "Testing.")
	If @error Then _ERROR($oDoc, "Failed to insert String. Error:" & @error & " Extended:" & @extended)

	; Retrieve the first Row's Object. Remember L.O. Rows are 0 based.
	$oRow = _LOCalc_RangeRowGetObjByPosition($oSheet, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Row Object. Error:" & @error & " Extended:" & @extended)

	; Set first Row's Height to optimal.
	_LOCalc_RangeRowHeight($oRow, True)
	If @error Then _ERROR($oDoc, "Failed to set Row Height. Error:" & @error & " Extended:" & @extended)

	; Select the word "Testing".
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_LEFT, 8, True)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Set the font weight to Bold.
	_LOCalc_TextCursorFont($oDoc, $oTextCursor, Null, Null, Null, $LOC_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to set text formatting. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of Paragraph objects
	$aoPar = _LOCalc_TextCursorParObjCreateList($oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of paragraph Objects. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I retrieved an Array of Paragraph Objects for Cell A1, there were " & UBound($aoPar) & " Paragraphs returned." & @CRLF & _
			"I will now set the first Paragraph Object's font settings to use 16 point Arial font type.")

	; Set the first paragraph Object's font to Arial, and font size to 16.
	_LOCalc_TextCursorFont($oDoc, $aoPar[0], "Arial", 16)
	If @error Then _ERROR($oDoc, "Failed to set text formatting. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
