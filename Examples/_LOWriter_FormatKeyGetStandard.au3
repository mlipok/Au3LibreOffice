#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $iFormatKey

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Query the standard Format Key for the Format Key type of $LOC_FORMAT_KEYS_DURATION
	$iFormatKey = _LOWriter_FormatKeyGetStandard($oDoc, $LOW_FORMAT_KEYS_DURATION)
	If @error Then _ERROR($oDoc, "Failed to retrieve the standard format key. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Standard format key number for format key type $LOC_FORMAT_KEYS_DURATION is: " & $iFormatKey & " It looks like this: " & _
			_LOWriter_FormatKeyGetString($oDoc, $iFormatKey))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
