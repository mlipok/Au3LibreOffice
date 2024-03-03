#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oTextCursor
	Local $avSettings[0]

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

	; Insert a Word
	_LOCalc_TextCursorInsertString($oTextCursor, "Hi! ")
	If @error Then _ERROR($oDoc, "Failed to insert String. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now set the word ""Hi!"" to Bold and Italic Arial font, size 16.")

	; Select the word Hi!
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GOTO_START, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Set the Word Hi to Font = "Arial", 16 point font, Italic and Bold
	_LOCalc_TextCursorFont($oDoc, $oTextCursor, "Arial", 16, $LOC_POSTURE_ITALIC, $LOC_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to set text font. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now set the font to Stencil, no Bold or Italic and font size 14, and then insert some text.")

	; Move the cursor to the end
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GOTO_END, 1, False)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Set the Font to = "Stencil", 14 point font, No Italic or Bold
	_LOCalc_TextCursorFont($oDoc, $oTextCursor, "Stencil", 14, $LOC_POSTURE_NONE, $LOC_WEIGHT_NORMAL)
	If @error Then _ERROR($oDoc, "Failed to set text font. Error:" & @error & " Extended:" & @extended)

	; Insert a Word
	_LOCalc_TextCursorInsertString($oTextCursor, " The End :)")
	If @error Then _ERROR($oDoc, "Failed to insert String. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current font settings. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_TextCursorFont($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve current format settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Font settings at the Cursor's current position are as follows: " & @CRLF & _
			"The Font name is: " & $avSettings[0] & @CRLF & _
			"The Font size is: " & $avSettings[1] & @CRLF & _
			"The current Italic setting is (See UDF Constants): " & $avSettings[2] & @CRLF & _
			"The current Bold setting is (See UDF Constants): " & $avSettings[3] & @CRLF & _
			"I will now demonstrate the return values when multiple settings are present in a selection.")

	; Move the cursor to the start, selecting all the words
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GOTO_START, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current font settings. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_TextCursorFont($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve current format settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Font settings at the Cursor's current position are as follows: " & @CRLF & _
			"The Font name is: " & $avSettings[0] & @CRLF & _
			"The Font size is: " & $avSettings[1] & @CRLF & _
			"The current Italic setting is (See UDF Constants): " & $avSettings[2] & @CRLF & _
			"The current Bold setting is (See UDF Constants): " & $avSettings[3] & @CRLF & _
			"Notice the values are weird when multiple font settings are present in a selection.")

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
