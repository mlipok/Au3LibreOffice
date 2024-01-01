#include <MsgBoxConstants.au3>
#include <Array.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell
	Local $asUndo[0], $asRedo[0]
	Local $iCount = 0

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Fill the Cell Range of A1 to C3 with numbers, one cell at a time. (Remember Columns and Rows are 0 based.)
	For $i = 0 To 2
		For $j = 0 To 2
			; Retrieve the Cell Object
			$oCell = _LOCalc_RangeGetCellByPosition($oSheet, $i, $j)
			If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

			; Set the Cell to a Number
			_LOCalc_CellValue($oCell, $iCount)
			If @error Then _ERROR($oDoc, "Failed to set Cell Value. Error:" & @error & " Extended:" & @extended)

			$iCount += 1

		Next

	Next

	; Undo one action.
	_LOCalc_DocUndo($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Undo the last action. Error:" & @error & " Extended:" & @extended)

	; Begin a Undo Action Group record. Name it "AutoIt Fill Cells"
	_LOCalc_DocUndoActionBegin($oDoc, "AutoIt Fill Cells")
	If @error Then _ERROR($oDoc, "Failed to begin an Undo Group record. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of available undo action titles.
	$asUndo = _LOCalc_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of available Redo action titles.
	$asRedo = _LOCalc_DocRedoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Redo action titles. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Here is a list of the available Undo actions, and the available Redo actions. Notice I also started an Undo Action group, which is listed in the Undo list.")

	; Display the available Undo action titles.
	_ArrayDisplay($asUndo)

	; Display the available Redo action titles.
	_ArrayDisplay($asRedo)

	; Clear the Undo/Redo list.
	_LOCalc_DocUndoReset($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Reset undo/redo action titles. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have reset the Undo/Redo Actions lists. I will retrieve the available Undo and Redo Actions lists again and show that they are now empty, including my Undo action group.")

	; Retrieve an array of available undo action titles again.
	$asUndo = _LOCalc_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of available Redo action titles again.
	$asRedo = _LOCalc_DocRedoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Redo action titles. Error:" & @error & " Extended:" & @extended)

	; Display the available Undo action titles again, if any.
	_ArrayDisplay($asUndo)

	; Display the available Redo action titles.
	_ArrayDisplay($asRedo)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
