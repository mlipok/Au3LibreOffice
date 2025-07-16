#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oFrameStyle, $oViewCursor, $oFrame
	Local $iMicrometers
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Frame Style named "Test Style"
	$oFrameStyle = _LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Frame into the document for demonstration.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR($oDoc, "Failed to create a Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Frame's style to my created style, "Test Style"
	_LOWriter_FrameStyleSet($oDoc, $oFrame, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to set Frame style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Frame Style Column count to 4.
	_LOWriter_FrameStyleColumnSettings($oFrameStyle, 4)
	If @error Then _ERROR($oDoc, "Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.25)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Frame Style Column size settings for column 2, set auto width to True, and Global spacing to 1/4".
	_LOWriter_FrameStyleColumnSize($oFrameStyle, 2, True, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleColumnSize($oFrameStyle, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame style's current Column size settings are as follows: " & @CRLF & _
			"Is Column width automatically adjusted? True/False: " & $avSettings[0] & @CRLF & _
			"The Global Spacing value for the entire frame, in Micrometers (If there is one): " & $avSettings[1] & @CRLF & _
			"The Spacing value between this column and the next column to the right is, in Micrometers: " & $avSettings[2] & @CRLF & _
			"The width of this column, in Micrometers: " & $avSettings[3] & @CRLF & _
			"Note: This value will be different from the UI value, even when converted to Inches or Centimeters, because the returned width value is a " & _
			"relative width, not a metric width, which is why I don't know how to set this value appropriately." & @CRLF & @CRLF & _
			"I will now demonstrate values when AutoWidth is deactivated.")

	; Set the Frame Style Column size settings for column 2, set auto width to False.
	_LOWriter_FrameStyleColumnSize($oFrameStyle, 2, False)
	If @error Then _ERROR($oDoc, "Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleColumnSize($oFrameStyle, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame style's new Column size settings are as follows: " & @CRLF & _
			"Is Column width automatically adjusted? True/False: " & $avSettings[0] & @CRLF & _
			"The Global Spacing value for the entire frame, in Micrometers (If there is one): " & $avSettings[1] & @CRLF & _
			"The Spacing value between this column and the next column to the right is, in Micrometers: " & $avSettings[2] & @CRLF & _
			"The width of this column, in Micrometers: " & $avSettings[3] & @CRLF & _
			"Note: This value is still different from the UI (even when converted) because, as I mentioned, the returned width value " & _
			"is a relative width, not a metric width value.")

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
