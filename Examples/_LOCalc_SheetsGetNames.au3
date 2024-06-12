#include <MsgBoxConstants.au3>
#include <Array.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2, $oSheet, $oCellRange
	Local $aavData[3]
	Local $avRowData[2]
	Local $sFilePathName, $sPath
	Local $asSheets[0]

	; Create a New, invisible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, True)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Create a new Sheet named "New Sheet".
	$oSheet = _LOCalc_SheetAdd($oDoc, "New Sheet")
	If @error Then _ERROR($oDoc, "Failed to create a new Sheet. Error:" & @error & " Extended:" & @extended)

	; Activate the new Sheet
	_LOCalc_SheetActivate($oDoc, $oSheet)
	If @error Then _ERROR($oDoc, "Failed to activate the Sheet. Error:" & @error & " Extended:" & @extended)

	; Fill my arrays with the desired Number and String Values I want in Column A to B.
	$avRowData[0] = "Seventy" ; A1
	$avRowData[1] = 5 ; B1
	$aavData[0] = $avRowData

	$avRowData[0] = 10 ; A2
	$avRowData[1] = 20 ; B2
	$aavData[1] = $avRowData

	$avRowData[0] = "Testing" ; A3
	$avRowData[1] = -1700 ; B3
	$aavData[2] = $avRowData

	; Retrieve Cell range A1 to B3
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended)

	$sFilePathName = @TempDir & "\TestExportDoc_" & @MDAY & ".ods"

	; Save The Document To Temp Directory.
	$sPath = _LOCalc_DocSaveAs($oDoc, $sFilePathName)
	If @error Then _ERROR($oDoc, "Failed to Save the Calc Document. Error:" & @error & " Extended:" & @extended)

	; Create a New, visible, Blank Libre Office Document.
	$oDoc2 = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended, Null, $sPath)

	; Link the Sheet "New Sheet" from Document 1 into this Document.
	_LOCalc_SheetLink($oDoc, $oDoc2, _LOCalc_SheetName($oDoc, $oSheet), $LOC_SHEET_LINK_MODE_NORMAL, True)
	If @error Then _ERROR($oDoc, "Failed to Link Calc Sheet. Error:" & @error & " Extended:" & @extended, $oDoc2, $sPath)

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended, $oDoc2, $sPath)

	; Add a new Sheet named "Test Sheet" after the first sheet.
	_LOCalc_SheetAdd($oDoc2, "Test Sheet", 1)
	If @error Then _ERROR($oDoc2, "Failed to Create a new Calc Sheet. Error:" & @error & " Extended:" & @extended, Null, $sPath)

	; Add a new Sheet Auto-named before the first sheet.
	_LOCalc_SheetAdd($oDoc2, Null, 0)
	If @error Then _ERROR($oDoc2, "Failed to Create a new Calc Sheet. Error:" & @error & " Extended:" & @extended, Null, $sPath)

	; Retrieve an Array of Sheet names.
	$asSheets = _LOCalc_SheetsGetNames($oDoc2)
	If @error Then _ERROR($oDoc2, "Failed to Retrieve an array of Sheet names. Error:" & @error & " Extended:" & @extended, Null, $sPath)

	_ArrayDisplay($asSheets)

	MsgBox($MB_OK, "", "I will now display a list of Linked Sheet names.")

	; Retrieve an Array of Linked Sheet names.
	$asSheets = _LOCalc_SheetsGetNames($oDoc2, True)
	If @error Then _ERROR($oDoc2, "Failed to Retrieve an array of Sheet names. Error:" & @error & " Extended:" & @extended, Null, $sPath)

	_ArrayDisplay($asSheets)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc2, False)
	If @error Then _ERROR($oDoc2, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended, Null, $sPath)

	; Delete the file.
	FileDelete($sPath)
EndFunc

Func _ERROR($oDoc, $sErrorText, $oDoc2 = Null, $sPath = Null)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	If IsObj($oDoc2) Then _LOCalc_DocClose($oDoc2, False)
	If IsString($sPath) Then FileDelete($sPath)
	Exit
EndFunc
