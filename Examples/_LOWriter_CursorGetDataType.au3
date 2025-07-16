#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iCursorDataType

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Place the Document Cursor in the document wherever you want to see what data type is returned, and press ok." & @CRLF & _
			"The possible cursor data type values are: " & @CRLF & _
			"$LOW_CURDATA_BODY_TEXT (1)" & @CRLF & _
			"$LOW_CURDATA_FRAME (2)" & @CRLF & _
			"$LOW_CURDATA_CELL (3)" & @CRLF & _
			"$LOW_CURDATA_FOOTNOTE (4)" & @CRLF & _
			"$LOW_CURDATA_ENDNOTE (5)" & @CRLF & _
			"$LOW_CURDATA_HEADER_FOOTER (6)")

	; Retrieve what type of Data the cursor is presently in.
	$iCursorDataType = _LOWriter_CursorGetDataType($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cursor Data type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	Switch $iCursorDataType
		Case $LOW_CURDATA_BODY_TEXT
			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cursor is currently in body text type of Data, with an integer value of : " & $iCursorDataType & _
					" — Or $LOW_CURDATA_BODY_TEXT")

		Case $LOW_CURDATA_FRAME
			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cursor is currently in Text Frame type of Data, with an integer value of : " & $iCursorDataType & _
					" — Or $LOW_CURDATA_FRAME")

		Case $LOW_CURDATA_CELL
			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cursor is currently in Text Table Cell type of Data, with an integer value of : " & $iCursorDataType & _
					" — Or $LOW_CURDATA_CELL")

		Case $LOW_CURDATA_FOOTNOTE
			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cursor is currently in Footnote type of Data, with an integer value of : " & $iCursorDataType & _
					" — Or $LOW_CURDATA_FOOTNOTE")

		Case $LOW_CURDATA_ENDNOTE
			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cursor is currently in Endnote type of Data, with an integer value of : " & $iCursorDataType & _
					" — Or $LOW_CURDATA_ENDNOTE")

		Case $LOW_CURDATA_HEADER_FOOTER
			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cursor is currently in Header/Footer type of Data, with an integer value of : " & $iCursorDataType & _
					" — Or $LOW_CURDATA_HEADER_FOOTER")

		Case Else
			MsgBox($MB_OK + $MB_TOPMOST, Default, "Something went wrong.")
	EndSwitch

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Following Error codes returned: Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
