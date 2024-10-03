#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $sPageStyleName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sPageStyleName = "NewPageStyle"

	MsgBox($MB_OK, "", "Does a Page Style Named """ & $sPageStyleName & """ exist in the document? True/False: " & _
			_LOCalc_PageStyleExists($oDoc, $sPageStyleName))

	; Create a New Page Style.
	_LOCalc_PageStyleCreate($oDoc, $sPageStyleName)
	If @error Then _ERROR($oDoc, "Failed to create a new Page Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK, "", "Now does a Page Style Named """ & $sPageStyleName & """exist in the document? True/False: " & _
			_LOCalc_PageStyleExists($oDoc, $sPageStyleName))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
