#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm1, $oForm2, $oSubForm
	Local $iCount

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm1 = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm2 = _LOWriter_FormAdd($oDoc, "AutoIt_Form2")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a sub-Form in the form.
	_LOWriter_FormAdd($oForm2, "AutoIt_SubForm")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create another sub-Form in the form.
	$oSubForm = _LOWriter_FormAdd($oForm2, "AutoIt_SubForm2")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a count of forms in the Document.
	$iCount = _LOWriter_FormsGetCount($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve a count of forms in the document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There is currently " & $iCount & " forms in the document. Press ok to move one sub-form to be a main form.")

	; Move one form to be a top-level form.
	_LOWriter_FormSubMove($oSubForm, $oDoc)
	If @error Then _ERROR($oDoc, "Failed to move form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a count of forms in the Document.
	$iCount = _LOWriter_FormsGetCount($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve a count of forms in the document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There are now " & $iCount & " forms in the document.")

	; Retrieve a count of sub-forms in Form 1.
	$iCount = _LOWriter_FormsGetCount($oForm1)
	If @error Then _ERROR($oDoc, "Failed to retrieve a count of sub-forms in the form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There is currently " & $iCount & " sub-forms in the AutoIt_Form. Press ok to move one top-level-form to be a sub-form of AutoIt_Form.")

	; Move one form from being a top-level form to be a sub-form.
	_LOWriter_FormSubMove($oForm2, $oForm1)
	If @error Then _ERROR($oDoc, "Failed to move form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a count of forms in Form 1.
	$iCount = _LOWriter_FormsGetCount($oForm1)
	If @error Then _ERROR($oDoc, "Failed to retrieve a count of sub-forms in the Form. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There is now " & $iCount & " sub-form in the AutoIt_Form.")

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
