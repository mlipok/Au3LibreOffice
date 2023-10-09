#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $iMicrometers, $iMicrometers2
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Frame into the document at the Viewcursor position, and 3000x3000 Micrometers wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	; Set Frame Border Width settings to: $LOW_BORDERWIDTH_MEDIUM on all four sides.
	_LOWriter_FrameBorderWidth($oFrame, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR("Failed to modify Frame settings. Error:" & @error & " Extended:" & @extended)

	; Convert 1/8" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.125)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Convert 1/4" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(.25)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Modify the Frame Border Padding Width settings to: 1/8" in all sides, and then 1/4" on the bottom.
	_LOWriter_FrameBorderPadding($oFrame, $iMicrometers, Null, $iMicrometers2)
	If @error Then _ERROR("Failed to set Frame settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameBorderPadding($oFrame)
	If @error Then _ERROR("Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame's Border Padding Width settings are as follows: " & @CRLF & _
			"The ""All"" Border Padding Width is, in Micrometers, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The Top Border Padding Width is, in Micrometers, (see UDF constants): " & $avSettings[1] & @CRLF & _
			"The Bottom Border Padding Width is, in Micrometers, (see UDF constants): " & $avSettings[2] & @CRLF & _
			"The Left Border Padding Width is, in Micrometers, (see UDF constants): " & $avSettings[3] & @CRLF & _
			"The Right Border Padding Width is, in Micrometers, (see UDF constants): " & $avSettings[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
