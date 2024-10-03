#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $iFormatKey
	Local $bExists

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a New Number Format Key.
	$iFormatKey = _LOCalc_FormatKeyCreate($oDoc, "#,##0.000")
	If @error Then _ERROR($oDoc, "Failed to create a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the new key exists, searching in Number Format Keys only.
	$bExists = _LOCalc_FormatKeyExists($oDoc, $iFormatKey, $LOC_FORMAT_KEYS_NUMBER)
	If @error Then _ERROR($oDoc, "Failed to search for a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If ($bExists = True) Then
		MsgBox($MB_OK, "", "I created a new Number format key. Its Key number is, " & $iFormatKey)
	Else
		MsgBox($MB_OK, "", "I Failed to create a new Number format key.")
	EndIf

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
