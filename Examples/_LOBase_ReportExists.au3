#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDoc, $oDBase, $oConnection
	Local $bReturn

	; Open the Libre Office Base Example Document.
	$oDoc = _LOBase_DocOpen(@ScriptDir & "\Extras\Example.odb")
	If @error Then Return _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; See if a Report exists with the name "rptReport1"
	$bReturn = _LOBase_ReportExists($oDoc, "rptReport1", True)
	If @error Then Return _ERROR($oDoc, "Failed to query if a Report exists by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a Report exist in the document with the name of ""rptReport1""? True/False. " & $bReturn)

	; See if a Report exists with the name "rptReport2" with Exhaustive set to True.
	$bReturn = _LOBase_ReportExists($oDoc, "rptReport2", True)
	If @error Then Return _ERROR($oDoc, "Failed to query if a Report exists by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a Report exist in the document with the name of ""rptReport2""? True/False. " & $bReturn)

	; See if a Report exists with the name "rptReport3" in the folder "Folder1".
	$bReturn = _LOBase_ReportExists($oDoc, "Folder1/rptReport3")
	If @error Then Return _ERROR($oDoc, "Failed to query if a Report exists by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a Report exist in the document with the name of ""rptReport3"" inside the folder ""Folder1""? True/False. " & $bReturn)

	; See if a Report exists with the name "rptReport3" with Exhaustive set to False.
	$bReturn = _LOBase_ReportExists($oDoc, "rptReport3", False)
	If @error Then Return _ERROR($oDoc, "Failed to query if a Report exists by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a Report exist in the document with the name of ""rptReport3"" with exhaustive set to False? True/False. " & $bReturn)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the Base document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
