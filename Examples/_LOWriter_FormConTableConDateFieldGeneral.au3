#include <MsgBoxConstants.au3>
#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oForm, $oControl, $oColumn
	Local $avColumn
	Local $iWidth
	Local $tDateMin, $tDateMax, $tDateDefault

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form in the Document.
	$oForm = _LOWriter_FormAdd($oDoc, "AutoIt_Form")
	If @error Then _ERROR($oDoc, "Failed to Create a form in the Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Form Control
	$oControl = _LOWriter_FormConInsert($oForm, $LOW_FORM_CON_TYPE_TABLE_CONTROL, 500, 300, 6000, 3000, "AutoIt_Form_Control")
	If @error Then _ERROR($oDoc, "Failed to insert a form control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Column
	$oColumn = _LOWriter_FormConTableConColumnAdd($oControl, $LOW_FORM_CON_TYPE_DATE_FIELD)
	If @error Then _ERROR($oDoc, "Failed to insert a Table control column. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1.25 inches to Micrometers
	$iWidth = _LOWriter_ConvertToMicrometer(1.25)
	If @error Then _ERROR($oDoc, "Failed to convert inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create the Minimum Date Structure.
	$tDateMin = _LOWriter_DateStructCreate(1992, 3, 28)
	If @error Then _ERROR($oDoc, "Failed to create a Date Structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create the Maximum Date Structure.
	$tDateMax = _LOWriter_DateStructCreate(2026, 7, 31)
	If @error Then _ERROR($oDoc, "Failed to create a Date Structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create the Default Date Structure.
	$tDateDefault = _LOWriter_DateStructCreate(2025, 5, 7)
	If @error Then _ERROR($oDoc, "Failed to create a Date Structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Column's General properties.
	_LOWriter_FormConTableConDateFieldGeneral($oColumn, "Renamed_AutoIt_Control", "Date Field 2", Null, True, True, False, $LOW_FORM_CON_MOUSE_SCROLL_FOCUS, $tDateMin, _
			$tDateMax, $LOW_FORM_CON_DATE_FRMT_SHORT_YYYYMMDD, $tDateDefault, True, True, 50, $iWidth, $LOW_ALIGN_HORI_LEFT, True, False, "Some Additional Information", _
			"This is Help Text", "www.HelpURL.fake")
	If @error Then _ERROR($oDoc, "Failed to set column's general properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the column. Return will be an Array in order of function parameters.
	$avColumn = _LOWriter_FormConTableConDateFieldGeneral($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Column's current settings are: " & @CRLF & _
			"The Column's name is: " & $avColumn[0] & @CRLF & _
			"The Column's Label is: " & $avColumn[1] & @CRLF & _
			"The Text Direction is: (See UDF Constants) " & $avColumn[2] & @CRLF & _
			"Is Strict formatting enabled? True/False: " & $avColumn[3] & @CRLF & _
			"Is the Column currently enabled? True/False: " & $avColumn[4] & @CRLF & _
			"Is the Column currently Read-Only? True/False: " & $avColumn[5] & @CRLF & _
			"What happens when the mouse scroll wheel is used over the Column (See UDF Constants): " & $avColumn[6] & @CRLF & _
			"This is the minimum Date structure. I'll just check if it is an Object. " & IsObj($avColumn[7]) & @CRLF & _
			"This is the maximum Date structure. I'll just check if it is an Object. " & IsObj($avColumn[8]) & @CRLF & _
			"What is date format? (See UDF Constants) " & $avColumn[9] & @CRLF & _
			"This is the Default Date structure. I'll just check if it is an Object. " & IsObj($avColumn[10]) & @CRLF & _
			"Does this Column act as a spin button? True/False: " & $avColumn[11] & @CRLF & _
			"Does the button action repeat if clicked and held? True/False: " & $avColumn[12] & @CRLF & _
			"If the button action repeats when clicked and held, how much delay is between each repeat? (In Milliseconds): " & $avColumn[13] & @CRLF & _
			"The Column's width is, in Micrometers: " & $avColumn[14] & @CRLF & _
			"The Horizontal Alignment is: (See UDF Constants) " & $avColumn[15] & @CRLF & _
			"Does this Column act as a drop down? True/False: " & $avColumn[16] & @CRLF & _
			"Will selections be hidden when losing focus? True/False: " & $avColumn[17] & @CRLF & _
			"The Additional Information text is: " & $avColumn[18] & @CRLF & _
			"The Help text is: " & $avColumn[19] & @CRLF & _
			"The Help URL is: " & $avColumn[20])

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
