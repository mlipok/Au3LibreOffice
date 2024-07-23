#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended)

	; Create another New, visible, Blank Libre Office Document.
	$oDoc2 = _LOBase_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended)

	$bReturn = _LOBase_DocIsActive($oDoc)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to query document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Document 1 the active document? True/False: " & $bReturn)

	$bReturn = _LOBase_DocIsActive($oDoc2)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to query document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Document 2 the active document? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close both documents.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	; Close the second document.
	_LOBase_DocClose($oDoc2, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $oDoc2, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsObj($oDoc2) Then _LOBase_DocClose($oDoc2, False)
	Exit
EndFunc
