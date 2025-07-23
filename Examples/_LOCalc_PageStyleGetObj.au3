#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a New Page Style.
	_LOCalc_PageStyleCreate($oDoc, "NewPageStyle")
	If @error Then _ERROR($oDoc, "Failed to create a new Page Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a Page Style Named ""NewPageStyle"" exist in the document? True/False: " & _
			_LOCalc_PageStyleExists($oDoc, "NewPageStyle"))

	; Retrieve the Page Style Object for the Page Style I just created named "NewPageStyle", so I can delete it now.
	$oPageStyle = _LOCalc_PageStyleGetObj($oDoc, "NewPageStyle")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the Page Style
	_LOCalc_PageStyleDelete($oDoc, $oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to delete the Page Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now does a Page Style Named ""NewPageStyle"" exist in the document? True/False: " & _
			_LOCalc_PageStyleExists($oDoc, "NewPageStyle"))

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
