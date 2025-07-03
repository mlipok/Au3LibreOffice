#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDoc, $oDBase, $oConnection
	Local $asReports[0], $asFolders[0]
	Local $sReports = "", $sFolders = ""

	; Open the Libre Office Base Example Document.
	$oDoc = _LOBase_DocOpen(@ScriptDir & "\Extras\Example.odb")
	If @error Then Return _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the Folder name exists already (This will be if a pevious example failed.) And delete it if so.
	If _LOBase_ReportFolderExists($oDoc, "Copied_Folder", False) Then _LOBase_ReportFolderDelete($oDoc, "Copied_Folder")
	If @error Then Return _ERROR($oDoc, "Failed to Check for pre-existing Report, or failed to delete it. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the Folder name exists already (This will be if a pevious example failed.) And delete it if so.
	If _LOBase_ReportFolderExists($oDoc, "Folder1/Copied_Folder2", False) Then _LOBase_ReportFolderDelete($oDoc, "Folder1/Copied_Folder2")
	If @error Then Return _ERROR($oDoc, "Failed to Check for pre-existing Report, or failed to delete it. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have a folder that contains a Report. Press ok to copy it and its contents.")

	; Copy the Folder.
	_LOBase_ReportFolderCopy($oDoc, "Folder1", "Copied_Folder")
	If @error Then Return _ERROR($oDoc, "Failed to copy the folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Copy the Folder again.
	_LOBase_ReportFolderCopy($oDoc, "Folder1", "Folder1/Copied_Folder2")
	If @error Then Return _ERROR($oDoc, "Failed to copy the folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Folder names.
	$asFolders = _LOBase_ReportFoldersGetNames($oDoc, True)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve array of Folder names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		$sFolders &= $asFolders[$i] & @CRLF
	Next

	; Retrieve an array of all the Reports contained in the Document.
	$asReports = _LOBase_ReportsGetNames($oDoc, True)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve an Array of Report names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		$sReports &= "- " & $asReports[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Here is a list of Folders contained in the document." & @CRLF & $sFolders & @CRLF & _
			"Here is a list of Reports contained in the document." & @CRLF & $sReports)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to close the document.")

	; Delete the Folder
	_LOBase_ReportFolderDelete($oDoc, "Copied_Folder")
	If @error Then Return _ERROR($oDoc, "Failed to delete a Report. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the second Folder
	_LOBase_ReportFolderDelete($oDoc, "Folder1/Copied_Folder2")
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
