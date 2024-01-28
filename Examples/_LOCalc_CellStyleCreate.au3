#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"


Example()

Func Example()
	Local $oDoc
	Local $bExists

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Create a new Cell Style to use for demonstration.
	_LOCalc_CellStyleCreate($oDoc, "NewCellStyle")
	If @error Then _ERROR($oDoc, "Failed to Create a new Cell Style. Error:" & @error & " Extended:" & @extended)

	; Check if the Cell style exists.
	$bExists = _LOCalc_CellStyleExists($oDoc, "NewCellStyle")
	If @error Then _ERROR($oDoc, "Failed to test for Cell Style existing in document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Cell style called ""NewCellStyle"" exist in the document? True/False: " & $bExists)

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
