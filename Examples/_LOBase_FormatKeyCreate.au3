#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDoc, $oReportDoc, $oDBase, $oConnection
	Local $iFormatKey

	; Open the Libre Office Base Example Document.
	$oDoc = _LOBase_DocOpen(@ScriptDir & "\Extras\Example.odb")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the Report name exists already (This will be if a pevious example failed.) And delete it if so.
	If _LOBase_ReportExists($oDoc, "rptAutoIt_Report", False) Then _LOBase_ReportDelete($oDoc, "rptAutoIt_Report")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to Check for pre-existing Report, or failed to delete it. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Make a copy of the contained Report.
	_LOBase_ReportCopy($oDoc, $oConnection, "rptReport1", "rptAutoIt_Report")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to copy a Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Report in Design Mode.
	$oReportDoc = _LOBase_ReportOpen($oDoc, $oConnection, "rptAutoIt_Report", True, True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to open a Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a New Number Format Key.
	$iFormatKey = _LOBase_FormatKeyCreate($oReportDoc, "#,##0.000")
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to create a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I created a new Number format key. Its Format Key number is: " & $iFormatKey & " It looks like this: " & _
			_LOBase_FormatKeyGetString($oReportDoc, $iFormatKey))

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the Report Document.
	_LOBase_ReportClose($oReportDoc, True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to close the Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the Report
	_LOBase_ReportDelete($oDoc, "rptAutoIt_Report")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to delete a Report. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

EndFunc   ;==>Example

Func _ERROR($oDoc, $oReportDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oReportDoc) Then _LOBase_ReportClose($oReportDoc, True)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
