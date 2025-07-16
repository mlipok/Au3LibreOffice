#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl
	Local $mFont
	Local $avControl

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_NAV_BAR, 500, 300, 6000, 2000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Font Descriptor.
	$mFont = _LOWriter_FontDescCreate("Times New Roman", $LOW_WEIGHT_BOLD, $LOW_POSTURE_ITALIC, 16, $LOW_COLOR_INDIGO, $LOW_UNDERLINE_BOLD, $LOW_COLOR_GREEN, $LOW_STRIKEOUT_NONE, False, $LOW_RELIEF_NONE)
	If @error Then _ERROR($oDoc, "Failed to create a Font Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's General properties.
	_LOWriter_FormConNavBarGeneral($oControl, "Renamed_AutoIt_Control", Null, True, True, True, 2, 50, $mFont, $LOW_COLOR_GOLD, $LOW_FORM_CON_BORDER_3D, False, True, True, True, True, "Some Additional Information", "This is Help Text", "www.HelpURL.fake")
	If @error Then _ERROR($oDoc, "Failed to modify the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the control. Return will be an Array in order of function parameters.
	$avControl = _LOWriter_FormConNavBarGeneral($oControl)
	If @error Then _ERROR($oDoc, "Failed to retrieve Control's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's current settings are: " & @CRLF & _
			"The Control's name is: " & $avControl[0] & @CRLF & _
			"The Text Direction is: (See UDF Constants) " & $avControl[1] & @CRLF & _
			"Is the Control currently enabled? True/False: " & $avControl[2] & @CRLF & _
			"Is the Control currently visible? True/False: " & $avControl[3] & @CRLF & _
			"Is the Control a Tab Stop position? True/False: " & $avControl[4] & @CRLF & _
			"If the Control is a Tab Stop position, what order position is it? " & $avControl[5] & @CRLF & _
			"The delay in Milliseconds between button repeats is: " & $avControl[6] & @CRLF & _
			"This is the current Font settings. I'll just check if it is a Map. " & IsMap($avControl[7]) & @CRLF & _
			"The Long Integer background color is: " & $avControl[8] & @CRLF & _
			"The Border Style is (See UDF Constants): " & $avControl[9] & @CRLF & _
			"Are Small icons used? True/False: " & $avControl[10] & @CRLF & _
			"Are positioning items shown? True/False: " & $avControl[11] & @CRLF & _
			"Are navigation items shown? True/False: " & $avControl[12] & @CRLF & _
			"Are action items shown? True/False: " & $avControl[13] & @CRLF & _
			"Are sorting and filtering items shown? True/False: " & $avControl[14] & @CRLF & _
			"The Additional Information text is: " & $avControl[15] & @CRLF & _
			"The Help text is: " & $avControl[16] & @CRLF & _
			"The Help URL is: " & $avControl[17])

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
