#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oMaster
	Local $sMasterFieldName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	$sMasterFieldName = "TestMaster"

	; Create a new Set Variable Master Field named "TestMaster".
	_LOWriter_FieldSetVarMasterCreate($oDoc, $sMasterFieldName)
	If @error Then _ERROR($oDoc, "Failed to create a Set Variable Master. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to delete the newly created Set Variable Master Field.")

	; Retrieve the Set Var. MasterField Object to use for deleting it.
	$oMaster = _LOWriter_FieldSetVarMasterGetObj($oDoc, $sMasterFieldName)
	If @error Then _ERROR($oDoc, "Failed to retrieve a Set Variable Master Object. Error:" & @error & " Extended:" & @extended)

	; Delete the Set Var. MasterField.
	_LOWriter_FieldSetVarMasterDelete($oDoc, $oMaster)
	If @error Then _ERROR($oDoc, "Failed to delete Master Field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does the Set Var Master Field still exist? True/False: " & _LOWriter_FieldSetVarMasterExists($oDoc, $sMasterFieldName))

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
