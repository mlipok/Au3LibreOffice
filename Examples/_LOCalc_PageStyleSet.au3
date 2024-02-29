#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet
	Local $sPageStyle

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Active Sheet's Object
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Set the Page style to Report using the ViewCursor.
	_LOCalc_PageStyleSet($oDoc, $oSheet, "Report")
	If @error Then _ERROR($oDoc, "Failed to set the page style. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Page Style set for this Sheet
	$sPageStyle = _LOCalc_PageStyleSet($oDoc, $oSheet)
	If @error Then _ERROR($oDoc, "Failed to retrieve the current page style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Page Style used by this sheet is: " & $sPageStyle)

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
