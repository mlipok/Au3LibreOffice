#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl
	Local $sValue
	Local $asValues[3] = ["Test1", "Testing2", "Tested3"]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_COMBO_BOX, 500, 300, 6000, 2000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set some values for the control.
	_LOWriter_FormConComboBoxGeneral($oControl, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, $asValues)
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's Value.
	_LOWriter_FormConComboBoxValue($oControl, "Tested2")
	If @error Then _ERROR($oDoc, "Failed to modify the Control's value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Notice I set the Control's value to a value that doesn't exist in the list. I'll now set it to one that does exist.")

	; Modify the Control's Value.
	_LOWriter_FormConComboBoxValue($oControl, "Testing2")
	If @error Then _ERROR($oDoc, "Failed to modify the Control's value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Value of the control.
	$sValue = _LOWriter_FormConComboBoxValue($oControl)
	If @error Then _ERROR($oDoc, "Failed to retrieve Control's value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's current value is: " & $sValue)

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
