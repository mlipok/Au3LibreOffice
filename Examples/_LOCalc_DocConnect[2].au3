#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2
	Local $iUserChoice
	Local $sDocName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have created a blank L.O. Calc Doc, I will now Connect to it and use the new Object returned to close it.")

	; Connect to the Current Document.
	$oDoc2 = _LOCalc_DocConnect("", True)
	If (@error > 0) Or Not IsObj($oDoc2) Then _ERROR($oDoc, $oDoc2, "Failed to Connect to Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve Doc Name.
	$sDocName = _LOCalc_DocGetName($oDoc, False)
	If @error Then _ERROR($oDoc, $oDoc2, "Failed to retrieve Calc Document name. Error:" & @error & " Extended:" & @extended)

	$iUserChoice = MsgBox($MB_YESNO, "Close?", "I have connected to the Document with the following title: " & $sDocName & _
			@CRLF & "Would you like to close it now?")

	If ($iUserChoice = $IDYES) Then
		; Close the document.
		_LOCalc_DocClose($oDoc2, False)
		If @error Then _ERROR($oDoc, $oDoc2, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
	EndIf

EndFunc

Func _ERROR($oDoc, $oDoc2, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	If IsObj($oDoc2) Then _LOCalc_DocClose($oDoc2, False)
	Exit
EndFunc
