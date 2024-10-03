#include <MsgBoxConstants.au3>
#include <Array.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $sDefault, $sSearch
	Local $asPrinters

	MsgBox($MB_OK, "", "I will list your currently available printers")

	$asPrinters = _LOCalc_DocEnumPrintersAlt()
	If @error Then _ERROR("Error retrieving array of Printers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK, "", "There were " & @extended & " results.")

	_ArrayDisplay($asPrinters)

	MsgBox($MB_OK, "", "I will list your currently default printer next.")

	$sDefault = _LOCalc_DocEnumPrintersAlt("", True)
	If @error Then _ERROR("Error retrieving Default Printer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If ($sDefault = "") Then
		MsgBox($MB_OK, "", "You do not have a default printer.")
	Else
		MsgBox($MB_OK, "", "Your default printer is: " & $sDefault)
	EndIf

	If (MsgBox($MB_YESNO, "", "We will search for a specific printer next, would you like to enter a phrase to search for?") = $IDYES) Then
		$sSearch = InputBox("", "Enter a search term, if the name is not full and exact, use an asterisk (*), such as ""*PDF*""", "*PDF*")
	Else
		MsgBox($MB_OK, "", "Okay, I will search for ""*PDF*""")
		$sSearch = "*PDF*"
	EndIf

	$asPrinters = _LOCalc_DocEnumPrintersAlt($sSearch)
	If @error Then _ERROR("Error retrieving array of Printers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK, "", "There were: " & @extended & " results")

	_ArrayDisplay($asPrinters)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
