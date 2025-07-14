#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl, $oLabel
	Local $mFont
	Local $avControl

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_NUMERIC_FIELD, 500, 300, 6000, 2000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Label Form Control
	$oLabel = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_LABEL, 3500, 2300, 3000, 1000, "AutoIt_Form_Label_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Font Descriptor.
	$mFont = _LOWriter_FontDescCreate("Times New Roman", $LOW_WEIGHT_BOLD, $LOW_POSTURE_ITALIC, 16, $LOW_COLOR_INDIGO, $LOW_UNDERLINE_BOLD, $LOW_COLOR_GREEN, $LOW_STRIKEOUT_NONE, False, $LOW_RELIEF_NONE)
	If @error Then _ERROR($oDoc, "Failed to create a Font Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's General properties.
	_LOWriter_FormConNumericFieldGeneral($oControl, "Renamed_AutoIt_Control", $oLabel, Null, False, True, True, False, True, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS, True, 1, 1.25, 50.00, 2, 18.44, 2, False, True, True, 50, $mFont, $LOW_ALIGN_HORI_CENTER, $LOW_ALIGN_VERT_BOTTOM, $LOW_COLOR_GOLD, $LOW_FORM_CON_BORDER_3D, $LOW_COLOR_BLACK, False, "Some Additional Information", "This is Help Text", "www.HelpURL.fake")
	If @error Then _ERROR($oDoc, "Failed to modify the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the control. Return will be an Array in order of function parameters.
	$avControl = _LOWriter_FormConNumericFieldGeneral($oControl)
	If @error Then _ERROR($oDoc, "Failed to retrieve Control's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's current settings are: " & @CRLF & _
			"The Control's name is: " & $avControl[0] & @CRLF & _
			"If there is a Label Control set for this control, this will be its Object. I'll just check IsObj: " & IsObj($avControl[1]) & @CRLF & _
			"The Text Direction is: (See UDF Constants) " & $avControl[2] & @CRLF & _
			"Is formatting strict? True/False: " & $avControl[3] & @CRLF & _
			"Is the Control currently enabled? True/False: " & $avControl[4] & @CRLF & _
			"Is the Control currently visible? True/False: " & $avControl[5] & @CRLF & _
			"Is the Control currently Read-Only? True/False: " & $avControl[6] & @CRLF & _
			"Is the Control currently Printable? True/False: " & $avControl[7] & @CRLF & _
			"What happens when the mouse scroll wheel is used over the control (See UDF Constants): " & $avControl[8] & @CRLF & _
			"Is the Control a Tab Stop position? True/False: " & $avControl[9] & @CRLF & _
			"If the Control is a Tab Stop position, what order position is it? " & $avControl[10] & @CRLF & _
			"The minimum value allowed to be entered is: " & $avControl[11] & @CRLF & _
			"The maximum value allowed to be entered is: " & $avControl[12] & @CRLF & _
			"The increment to increase the value is: " & $avControl[13] & @CRLF & _
			"The default value is: " & $avControl[14] & @CRLF & _
			"How many decimal places are there?: " & $avControl[15] & @CRLF & _
			"Is a Thousands separator used? True/False: " & $avControl[16] & @CRLF & _
			"Does the field have a Spin button? True/False: " & $avControl[17] & @CRLF & _
			"If the spin button is clicked and held, does the action repeat? True/False: " & $avControl[18] & @CRLF & _
			"What is the delay, in Milliseconds, that the action repeats?: " & $avControl[19] & @CRLF & _
			"This is the current Font settings. I'll just check if it is a Map. " & IsMap($avControl[20]) & @CRLF & _
			"The Horizontal Alignment is: (See UDF Constants) " & $avControl[21] & @CRLF & _
			"The Vertical Alignment is: (See UDF Constants) " & $avControl[22] & @CRLF & _
			"The Long Integer background color is: " & $avControl[23] & @CRLF & _
			"The Border Style is (See UDF Constants): " & $avControl[24] & @CRLF & _
			"The Border color is, in Long Integer format: " & $avControl[25] & @CRLF & _
			"Are selections hidden when the control loses focus? True/False: " & $avControl[26] & @CRLF & _
			"The Additional Information text is: " & $avControl[27] & @CRLF & _
			"The Help text is: " & $avControl[28] & @CRLF & _
			"The Help URL is: " & $avControl[29])

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
