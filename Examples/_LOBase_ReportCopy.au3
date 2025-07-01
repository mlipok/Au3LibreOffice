#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDoc, $oDBase, $oConnection
	Local $asReports
	Local $sReports = ""

	; Open the Libre Office Base Example Document.
	$oDoc = _LOBase_DocOpen(@ScriptDir & "\Extras\Example.odb")
	If @error Then Return _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to copy some Reports.")

	; Make a copy of the contained Report.
	_LOBase_ReportCopy($oDoc, $oConnection, "rptReport1", "rptAutoIt_Report")
	If @error Then Return _ERROR($oDoc, "Failed to copy a Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Make a copy of the contained Report in Folder1.
	_LOBase_ReportCopy($oDoc, $oConnection, "rptReport1", "Folder1/rptAutoIt_Report4")
	If @error Then Return _ERROR($oDoc, "Failed to copy a Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Make a copy of the Report contained in Folder1 to the main level.
	_LOBase_ReportCopy($oDoc, $oConnection, "Folder1/rptReport2", "rptAutoIt_Report5")
	If @error Then Return _ERROR($oDoc, "Failed to copy a Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of all the Reports contained in the Document.
	$asReports = _LOBase_ReportsGetNames($oDoc, True)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve an Array of Report names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		$sReports &= "- " & $asReports[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Document contains the following Reports:" & @CRLF & $sReports)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to close the document.")

	; Delete the Report
	_LOBase_ReportDelete($oDoc, "rptAutoIt_Report")
	If @error Then Return _ERROR($oDoc, "Failed to delete a Report. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the second Report
	_LOBase_ReportDelete($oDoc, "Folder1/rptAutoIt_Report4")
	If @error Then Return _ERROR($oDoc, "Failed to delete a Report. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the third Report
	_LOBase_ReportDelete($oDoc, "rptAutoIt_Report5")
	If @error Then Return _ERROR($oDoc, "Failed to delete a Report. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	Exit
EndFunc
