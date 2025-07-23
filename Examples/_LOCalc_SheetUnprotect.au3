#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet
	Local $bProtected

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the currently Active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Retrieve the currently Active Sheet's Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Protect the current sheet with the password 1234
	_LOCalc_SheetProtect($oSheet, "1234")
	If @error Then _ERROR($oDoc, "Failed to password protect the current sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check again if the Sheet is currently protected
	$bProtected = _LOCalc_SheetIsProtected($oSheet)
	If @error Then _ERROR($oDoc, "Failed to check if Sheet is currently password protected. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the currently active Sheet protected? True/False: " & $bProtected & @CRLF & _
			"Press ok to unprotect the Sheet.")

	; Unprotect the Sheet.
	_LOCalc_SheetUnprotect($oSheet, "1234")
	If @error Then _ERROR($oDoc, "Failed to remove password protection from the current sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check again if the Sheet is currently protected
	$bProtected = _LOCalc_SheetIsProtected($oSheet)
	If @error Then _ERROR($oDoc, "Failed to check if Sheet is currently password protected. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the currently active Sheet protected? True/False: " & $bProtected)

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
