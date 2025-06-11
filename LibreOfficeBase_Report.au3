#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;~ #Tidy_Parameters=/sf /reel
#include-once
#include "LibreOffice_Constants.au3"

; Common includes for Base
#include "LibreOfficeBase_Internal.au3"

; Other includes for Base

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Adding, Deleting, and modifying, etc. L.O. Base Reports.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; Notes .........: Presently I am unable to create a new Report.
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_ReportClose
; _LOBase_ReportConnect
; _LOBase_ReportDelete
; _LOBase_ReportExists
; _LOBase_ReportFolderCreate
; _LOBase_ReportFolderDelete
; _LOBase_ReportFolderExists
; _LOBase_ReportFolderRename
; _LOBase_ReportFoldersGetCount
; _LOBase_ReportFoldersGetNames
; _LOBase_ReportIsModified
; _LOBase_ReportOpen
; _LOBase_ReportRename
; _LOBase_ReportSave
; _LOBase_ReportsGetCount
; _LOBase_ReportsGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportClose
; Description ...: Close an opened Report Document.
; Syntax ........: _LOBase_ReportClose(ByRef $oReportDoc[, $bForceClose = False])
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportOpen, _LOBase_ReportConnect, or _LOBase_ReportCreate function.
;                  $bForceClose         - [optional] a boolean value. Default is False. If True, the Report document will be closed regardless if there are unsaved changes. See remarks.
; Return values .: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bForceClose not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Document has been modified and not saved, and $bForceClose is False.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify Report Document's name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 7 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  @Error 3 @Extended 8 Return 0 = Failed to identify Report in Parent Document.
;                  @Error 3 @Extended 9 Return 0 = Document called in $oReportDoc not a Report Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning a Boolean value of whether the Report Document was successfully closed (True), or not.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If there are unsaved changes in the document when close is called, and $bForceClose is True, they will be lost.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportClose(ByRef $oReportDoc, $bForceClose = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn
	Local $oReport, $oSource, $oObj
	Local $sTitle
	Local $iCount = 0, $iFolders = 1
	Local $asResult[0], $asNames[0], $asFolderList[0]
	Local $avFolders[0][2]
	Local Enum $iName, $iObj, $iPrefix
	Local Const $__STR_REGEXPARRAYMATCH = 1

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bForceClose) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oReportDoc.isModified() And Not $bForceClose Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oReportDoc.supportsService("com.sun.star.text.TextDocument") Then ; Report Doc is in viewing/Read-Only mode.
		$oReportDoc.close(True)
		$bReturn = True

	ElseIf $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then  ; Report is in Design mode.
		; "Testing.odb : abc1Report2"
		$asResult = StringRegExp($oReportDoc.Title(), "\: (.+)$", $__STR_REGEXPARRAYMATCH) ; Retrieve Report Title.
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$sTitle = $asResult[0]

		$oSource = $oReportDoc.Parent.ReportDocuments()
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$asNames = $oSource.getElementNames()
		If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oSource.hasByName($sTitle) And $oSource.getByName($sTitle).supportsService("com.sun.star.ucb.Content") And ($oSource.getByName($sTitle).getComponent() = $oReportDoc) Then ; getComponent will either return a Object or not (If the report isn't currently open.). If it does, and is the same as the called $oReportDoc, it is the one to close.
			$oReport = $oSource.getByName($sTitle)

		Else

			For $i = 0 To UBound($asNames) - 1
				$oObj = $oSource.getByName($asNames[$i])
				If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
					ReDim $avFolders[1][2]
					$avFolders[0][$iName] = $asNames[$i]
					$avFolders[0][$iObj] = $oObj
				EndIf

				While ($iCount < UBound($avFolders))
					If $avFolders[$iCount][$iObj].hasByName($sTitle) And $avFolders[$iCount][$iObj].getByName($sTitle).supportsService("com.sun.star.ucb.Content") And _
							($avFolders[$iCount][$iObj].getByName($sTitle).getComponent() = $oReportDoc) Then ; getComponent will either return a Object or not (If the report isn't currently open.). If it does, and is the same as the called $oReportDoc, it is the one to close.
						$oReport = $avFolders[$iCount][$iObj].getByName($sTitle)
						ExitLoop
					EndIf

					$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
					If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					For $k = 0 To UBound($asFolderList) - 1
						$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
						If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

						If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
							ReDim $avFolders[$iFolders + 1][2]
							$avFolders[$iFolders][$iName] = $asFolderList[$k]
							$avFolders[$iFolders][$iObj] = $oObj

							$iFolders += 1
						EndIf
					Next

					$iCount += 1
				WEnd

				If (UBound($avFolders) > 0) Then ReDim $avFolders[0][2]
				$iCount = 0
				$iFolders = 1
			Next
		EndIf

		If Not IsObj($oReport) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

		If $oReportDoc.isModified() Then $oReportDoc.Modified = False ; Set modified to false, so the user wont be prompted.

		$bReturn = $oReport.Close()

	Else ; Error, unknown document?
		Return SetError($__LO_STATUS_PROCESSING_ERROR, 9, 0)

	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_ReportClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConnect
; Description ...: Retrieve an Object for the currently open Report or Reports.
; Syntax ........: _LOBase_ReportConnect([$bConnectCurrent = True])
; Parameters ....: $bConnectCurrent     - [optional] a boolean value. Default is True. If True, Returns an Object for the last active Report. Else an array of all Open Reports. See Remarks.
; Return values .: Success: Object or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bConnectCurrent not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.ServiceManager Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create com.sun.star.frame.Desktop Object.
;                  @Error 2 @Extended 3 Return 0 = Failed to create enumeration of open Documents.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = No LibreOffice windows are open.
;                  @Error 3 @Extended 2 Return 0 = Current LibreOffice window is not a Report Document.
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Success. Connected to the currently active window, returning the Report Document Object. Report is in Read-Only viewing mode.
;                  @Error 0 @Extended 2 Return Object = Success. Connected to the currently active window, returning the Report Document Object. Report is in Design mode.
;                  @Error 0 @Extended ? Return Array = Success. Returning a Three columned Array with all open Report Documents. @Extended is set to the number of results. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Returned array when connecting to all open Report Documents returns an array with Three columns per result. ($aArray[0][3]). Each result is stored in a separate row;
;                  Row 1, Column 0 contain the Object for that document. e.g. $aArray[0][0] = $oDoc
;                  Row 1, Column 1 contains the Document's full title with extension and the Report Name, separated by a colon. e.g. $aArray[0][1] = "Testing.odb : Report1"
;                  Row 1, Column 2 contains a Boolean value whether the Report is in Design mode (True) or not.
;                  Row 2, Column 0 contain the Object for the next document. And so on. e.g. $aArray[1][0] = $oDoc2
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConnect($bConnectCurrent = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local $aoConnectAll[0][3]
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop
	Local $sReportViewServiceName = "com.sun.star.text.TextDocument", $sReportDesignServiceName = "com.sun.star.report.ReportDefinition"

	If Not IsBool($bConnectCurrent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not $oDesktop.getComponents.hasElements() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; no L.O open
	$oEnumDoc = $oDesktop.getComponents.createEnumeration()
	If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If $bConnectCurrent Then
		$oDoc = $oDesktop.currentComponent()
		If ($oDoc.supportsService($sReportViewServiceName) And $oDoc.isReadOnly() And Not (IsObj($oDoc.Parent()))) Then ; View only Report Doc.

			Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc)

		ElseIf $oDoc.supportsService($sReportDesignServiceName) Then ; Report Doc in Design mode.

			Return SetError($__LO_STATUS_SUCCESS, 2, $oDoc)

		Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndIf

	Else

		ReDim $aoConnectAll[1][3]
		$iCount = 0
		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService($sReportDesignServiceName) _ ; Report Doc in Design mode.
					Or ($oDoc.supportsService($sReportViewServiceName) And $oDoc.isReadOnly() And Not (IsObj($oDoc.Parent()))) Then ; If Parent is not present and document is Read-Only, it should be a Database Report.

				ReDim $aoConnectAll[$iCount + 1][3]
				$aoConnectAll[$iCount][0] = $oDoc
				$aoConnectAll[$iCount][1] = $oDoc.Title()
				$aoConnectAll[$iCount][2] = ($oDoc.supportsService($sReportDesignServiceName)) ? (True) : (False)
				$iCount += 1
			EndIf
			Sleep(10)
		WEnd

		Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoConnectAll)
	EndIf
EndFunc   ;==>_LOBase_ReportConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportDelete
; Description ...: Delete a Report from a Document.
; Syntax ........: _LOBase_ReportDelete(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sName               - a string value. The Report name to Delete. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Folder or Sub-Folder not found.
;                  @Error 1 @Extended 4 Return 0 = Name called in $sName not found in Folder.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sName not a Report.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete Report.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Report was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To delete a report contained in a folder, you MUST prefix the Report name called in $sName by the folder path it is located in, separated by forward slashes (/). e.g. to delete ReportXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/ReportXYZ
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportDelete(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Report name to delete

	If Not $oSource.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not $oSource.getByName($sName).supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSource.removeByName($sName)

	If $oSource.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportExists
; Description ...: Check whether a Document contains a Report by name.
; Syntax ........: _LOBase_ReportExists(ByRef $oDoc, $sName[, $bExhaustive = True])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sName               - a string value. The name of the Report to look for. See remarks.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, the search looks inside sub-folders.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended ? Return Boolean = Success. Returning a Boolean value indicating if the Document contains a Report by the called name (True) or not. If True, and $bExhaustive is True, @Extended is set to the number of times a Report with the same name is found in the Document (In sub-folders).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To narrow the search for a Report down to a specific folder, you MUST prefix the Report name called in $sName by the folder path to look in, separated by forward slashes (/). e.g. to search for ReportXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/ReportXYZ
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportExists(ByRef $oDoc, $sName, $bExhaustive = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local $bReturn = False
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iReports = 0
	Local $asNames[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Report name to search

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If $oSource.hasByName($sName) And $oSource.getByName($sName).supportsService("com.sun.star.ucb.Content") Then
		$iReports += 1
		$bReturn = True
	EndIf

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			ReDim $avFolders[1][2]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				If $avFolders[$iCount][$iObj].hasByName($sName) And $avFolders[$iCount][$iObj].getByName($sName).supportsService("com.sun.star.ucb.Content") Then
					$iReports += 1
					$bReturn = True
				EndIf

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						ReDim $avFolders[$iFolders + 1][2]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][2]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iReports, $bReturn)
EndFunc   ;==>_LOBase_ReportExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFolderCreate
; Description ...: Create a Report Folder.
; Syntax ........: _LOBase_ReportFolderCreate(ByRef $oDoc, $sFolder)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sFolder             - a string value. The Folder name to create. Can also include the sub-folder path. See Remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 3 Return 0 = Folder or Sub-Folder not found.
;                  @Error 1 @Extended 4 Return 0 = Name called in $sFolder already exists in Folder.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to insert new Folder into Base Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully created a Folder.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To create a Folder inside a folder, the Folder name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to create FolderXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3/FolderXYZ.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFolderCreate(ByRef $oDoc, $sFolder)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oObj
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sFolder, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sFolder = $asSplit[$asSplit[0]] ; Last element of Array will be the Folder name to Create

	If $oSource.hasByName($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oObj = $oSource.createInstance("com.sun.star.sdb.Reports")

	$oSource.insertbyName($sFolder, $oObj)

	If Not $oSource.hasByName($sFolder) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportFolderCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFolderDelete
; Description ...: Delete a Report Folder from a Document.
; Syntax ........: _LOBase_ReportFolderDelete(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sName               - a string value. The Folder name to Delete. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Folder or Sub-Folder not found.
;                  @Error 1 @Extended 4 Return 0 = Name called in $sName not found in Folder.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sName not a Folder.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Folder was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To delete a Folder contained in a folder, you MUST prefix the Folder name called in $sName by the folder path it is located in, separated by forward slashes (/). e.g. to delete FolderXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/FolderXYZ
;                  Deleting a Folder will delete all contents also.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFolderDelete(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Folder name to Delete

	If Not $oSource.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not $oSource.getByName($sName).supportsService("com.sun.star.sdb.Reports") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSource.removeByName($sName)

	If $oSource.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportFolderDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFolderExists
; Description ...: Check whether a Document contains a Report Folder by name.
; Syntax ........: _LOBase_ReportFolderExists(ByRef $oDoc, $sName[, $bExhaustive = True])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sName               - a string value. The name of the Folder to look for.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, the search looks inside sub-folders.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended ? Return Boolean = Success. Returning a Boolean value indicating if the Document contains a Folder by the called name (True) or not. If True, and $bExhaustive is True, @Extended is set to the number of times a Folder with the same name is found in the Document (In sub-folders).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To narrow the search for a Folder down to a specific folder, you MUST prefix the Folder name called in $sName by the folder path to look in, separated by forward slashes (/). e.g. to search for FolderXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/FolderXYZ
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFolderExists(ByRef $oDoc, $sName, $bExhaustive = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local $bReturn = False
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iResults = 0
	Local $asNames[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Folder name to search

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If $oSource.hasByName($sName) And $oSource.getByName($sName).supportsService("com.sun.star.sdb.Reports") Then
		$iResults += 1
		$bReturn = True
	EndIf

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			ReDim $avFolders[1][2]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				If $avFolders[$iCount][$iObj].hasByName($sName) And $avFolders[$iCount][$iObj].getByName($sName).supportsService("com.sun.star.sdb.Reports") Then
					$iResults += 1
					$bReturn = True
				EndIf

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						ReDim $avFolders[$iFolders + 1][2]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][2]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iResults, $bReturn)
EndFunc   ;==>_LOBase_ReportFolderExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFolderRename
; Description ...: Rename a Report Folder.
; Syntax ........: _LOBase_ReportFolderRename(ByRef $oDoc, $sFolder, $sNewName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sFolder             - a string value. The Folder to rename, including the Sub-Folder path, if applicable. See Remarks.
;                  $sNewName            - a string value. The New name to rename the Report Folder to.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 3 Return 0 = $sNewName not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder not found.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sNewName already exists in Folder.
;                  @Error 1 @Extended 6 Return 0 = Folder name called in $sFolder not found in Folder or is not a Folder.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to rename folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully renamed the Folder
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To rename a Folder inside a folder, the original Folder name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to rename FolderXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3/FolderXYZ.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFolderRename(ByRef $oDoc, $sFolder, $sNewName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sFolder, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sFolder = $asSplit[$asSplit[0]] ; Last element of Array will be the Folder name to Rename

	If $oSource.hasByName($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If Not $oSource.hasByName($sFolder) Or Not $oSource.getByName($sFolder).supportsService("com.sun.star.sdb.Reports") Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oSource.getByName($sFolder).rename($sNewName)

	If Not $oSource.hasByName($sNewName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportFolderRename

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFoldersGetCount
; Description ...: Retrieve a count of Report Folders contained in the Document.
; Syntax ........: _LOBase_ReportFoldersGetCount(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, retrieves a count of all folders, including those in sub-folders.
;                  $sFolder             - [optional] a string value. Default is "". The Folder to return the count of folders for. See remarks.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder called in $sFolder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Report Folders contained in the Document as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only the count of main level Folders (not located in folders), or if $bExhaustive is set to True, it will return a count of all Folders contained in the document.
;                  You can narrow the Folder count down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get a count of Folders contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFoldersGetCount(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iResults = 0
	Local $asNames[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFolder <> "") Then
		$asSplit = StringSplit($sFolder, "/")

		For $i = 1 To $asSplit[0]
			If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
			$oSource = $oSource.getByName($asSplit[$i])
			If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Next
	EndIf

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			$iResults += 1
			ReDim $avFolders[1][2]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						$iResults += 1
						ReDim $avFolders[$iFolders + 1][2]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][2]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, $iResults)
EndFunc   ;==>_LOBase_ReportFoldersGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFoldersGetNames
; Description ...: Retrieve an array of Folder Names contained in a Document.
; Syntax ........: _LOBase_ReportFoldersGetNames(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, the search looks inside sub-folders.
;                  $sFolder             - [optional] a string value. Default is "". The Sub-Folder to return the array of Folder names from. See remarks.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder called in $sFolder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Folder names contained in this Document. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only an array of main level Folder names (not located in sub-folders), or if $bExhaustive is set to True, it will return an array of all folders contained in the document.
;                  You can narrow the Folder name return down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get an array of Folders contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
;                  All Folders located in sub-folders will have the folder path prefixed to the Folder name, separated by forward slashes (/). e.g. Folder1/Folder2/Folder3.
;                  Calling $bExhaustive with True when searching inside a Folder, will get all Folder names from inside that folder, and all sub-folders.
;                  The order of the Folder names inside the folders may not necessarily be in proper order, i.e. if there are two sub folders, and folders inside the first sub-folder, the two folders will be listed first, then the folders inside the first sub-folder.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFoldersGetNames(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iResults = 0
	Local $asNames[0], $asFolders[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFolder <> "") Then
		$asSplit = StringSplit($sFolder, "/")

		For $i = 1 To $asSplit[0]
			If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
			$oSource = $oSource.getByName($asSplit[$i])
			If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Next
		$sFolder &= "/"
	EndIf

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			If (UBound($asFolders) >= $iResults) Then ReDim $asFolders[$iResults + 1]
			$asFolders[$iResults] = $sFolder & $asNames[$i]
			$iResults += 1
			ReDim $avFolders[1][3]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
			$avFolders[0][$iPrefix] = $asNames[$i] & "/"
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						If (UBound($asFolders) >= $iResults) Then ReDim $asFolders[$iResults + 1]
						$asFolders[$iResults] = $sFolder & $avFolders[$iCount][$iPrefix] & $asFolderList[$k]
						$iResults += 1
						ReDim $avFolders[$iFolders + 1][3]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj
						$avFolders[$iFolders][$iPrefix] = $avFolders[$iCount][$iPrefix] & $asFolderList[$k] & "/"

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][3]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($asFolders), $asFolders)
EndFunc   ;==>_LOBase_ReportFoldersGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportIsModified
; Description ...: Test whether the Report has been modified since being created or since the last save.
; Syntax ........: _LOBase_ReportIsModified(ByRef $oReportDoc)
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportOpen, _LOBase_ReportConnect, or _LOBase_ReportCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the Report has been modified since last being saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportIsModified(ByRef $oReportDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oReportDoc.isModified())
EndFunc   ;==>_LOBase_ReportIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportOpen
; Description ...: Open a Report Document
; Syntax ........: _LOBase_ReportOpen(ByRef $oDoc, ByRef $oConnection, $sName[, $bDesign = True[, $bVisible = True]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sName               - a string value. The Report name to Open. See remarks.
;                  $bDesign             - [optional] a boolean value. Default is True. If True, the Report is opened in Design mode.
;                  $bVisible            - [optional] a boolean value. Default is True. If True, the Report document will be visible when opened.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $bDesign not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = Folder or Sub-Folder not found.
;                  @Error 1 @Extended 7 Return 0 = Name called in $sName not found in Folder.
;                  @Error 1 @Extended 8 Return 0 = Name called in $sName not a Report.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to open Report Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning opened Report Document's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To open a Report located inside a folder, the Report name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to open ReportXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/ReportXYZ.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportOpen(ByRef $oDoc, ByRef $oConnection, $sName, $bDesign = True, $bVisible = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource, $oReportDoc
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDesign) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If Not $oDoc.CurrentController.isConnected() Then $oDoc.CurrentController.connect()

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Report name to Open

	If Not $oSource.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$oObj = $oSource.getByName($sName)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	If Not $oObj.supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	If $bDesign Then
		$oReportDoc = $oObj.openDesign()

	Else
		$oReportDoc = $oObj.open()
	EndIf

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$oReportDoc.CurrentController.Frame.ContainerWindow.Visible = $bVisible
	If ($oReportDoc.CurrentController.Frame.ContainerWindow.isVisible() <> $bVisible) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, $oReportDoc)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oReportDoc)
EndFunc   ;==>_LOBase_ReportOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportRename
; Description ...: Rename a Report.
; Syntax ........: _LOBase_ReportRename(ByRef $oDoc, $sReport, $sNewName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sReport             - a string value. The Report to rename, including the Sub-Folder path, if applicable. See Remarks.
;                  $sNewName            - a string value. The New name to rename the Report to.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sReport not a String.
;                  @Error 1 @Extended 3 Return 0 = $sNewName not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder not found.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sNewName already exists in Folder.
;                  @Error 1 @Extended 6 Return 0 = Report name called in $sReport not found in Folder or is not a Report.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to rename Report.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully renamed the Report.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To rename a Report inside a folder, the original Report name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to rename ReportXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sReport with the following path: Folder1/Folder2/Folder3/ReportXYZ.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportRename(ByRef $oDoc, $sReport, $sNewName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sReport) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sReport, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sReport = $asSplit[$asSplit[0]] ; Last element of Array will be the Report name to Rename

	If $oSource.hasByName($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If Not $oSource.hasByName($sReport) Or Not $oSource.getByName($sReport).supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oSource.getByName($sReport).rename($sNewName)

	If Not $oSource.hasByName($sNewName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportRename

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportSave
; Description ...: Save any changes made to a Document.
; Syntax ........: _LOBase_ReportSave(ByRef $oReportDoc)
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportOpen, _LOBase_ReportConnect, or _LOBase_ReportCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Document called in $oReportDoc is read only.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Report Document's name.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  @Error 3 @Extended 7 Return 0 = Document called in $oReportDoc not a Report Document.
;                  @Error 3 @Extended 8 Return 0 = Failed to identify Report in Parent Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Report was successfully saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportSave(ByRef $oReportDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oReport, $oObj
	Local $asResult[0], $asNames[0], $asFolderList[0]
	Local $iCount = 0, $iFolders = 1
	Local $avFolders[0][2]
	Local $sTitle
	Local Enum $iName, $iObj, $iPrefix
	Local Const $__STR_REGEXPARRAYMATCH = 1

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oReportDoc.supportsService("com.sun.star.text.TextDocument") And $oReportDoc.isReadOnly() Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Nothing to save in a Read only Doc.

	If $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then ; Report is in Design mode.
		; "Testing.odb : abc1Report2"
		$asResult = StringRegExp($oReportDoc.Title(), "\: (.+)$", $__STR_REGEXPARRAYMATCH) ; Retrieve Report Title.
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$sTitle = $asResult[0]

		$oSource = $oReportDoc.Parent.ReportDocuments()
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$asNames = $oSource.getElementNames()
		If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		If $oSource.hasByName($sTitle) And $oSource.getByName($sTitle).supportsService("com.sun.star.ucb.Content") And ($oSource.getByName($sTitle).getComponent() = $oReportDoc) Then ; getComponent will either return a Object or not (If the report isn't currently open.). If it does, and is the same as the called $oReportDoc, it is the one to close.
			$oReport = $oSource.getByName($sTitle)

		Else

			For $i = 0 To UBound($asNames) - 1
				$oObj = $oSource.getByName($asNames[$i])
				If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

				If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
					ReDim $avFolders[1][2]
					$avFolders[0][$iName] = $asNames[$i]
					$avFolders[0][$iObj] = $oObj
				EndIf

				While ($iCount < UBound($avFolders))
					If $avFolders[$iCount][$iObj].hasByName($sTitle) And $avFolders[$iCount][$iObj].getByName($sTitle).supportsService("com.sun.star.ucb.Content") And _
							($avFolders[$iCount][$iObj].getByName($sTitle).getComponent() = $oReportDoc) Then ; getComponent will either return a Object or not (If the report isn't currently open.). If it does, and is the same as the called $oReportDoc, it is the one to close.
						$oReport = $avFolders[$iCount][$iObj].getByName($sTitle)
						ExitLoop
					EndIf

					$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
					If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

					For $k = 0 To UBound($asFolderList) - 1
						$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
						If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

						If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
							ReDim $avFolders[$iFolders + 1][2]
							$avFolders[$iFolders][$iName] = $asFolderList[$k]
							$avFolders[$iFolders][$iObj] = $oObj

							$iFolders += 1
						EndIf
					Next

					$iCount += 1
				WEnd

				If (UBound($avFolders) > 0) Then ReDim $avFolders[0][2]
				$iCount = 0
				$iFolders = 1
			Next
		EndIf

	Else ; Error, unknown document?
		Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

	EndIf

	If Not IsObj($oReport) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

	$oReport.Store()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportsGetCount
; Description ...: Retrieve a count of Reports contained in the Document.
; Syntax ........: _LOBase_ReportsGetCount(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, retrieves a count of all Reports, including those in sub-folders.
;                  $sFolder             - [optional] a string value. Default is "". The Folder to return the count of Reports for. See remarks.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder called in $sFolder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Reports contained in the Document, as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only the count of main level Reports (not located in folders), or if $bExhaustive is set to True, the return will be a count of all Reports contained in the document.
;                  You can narrow the Report count down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get a count of Reports contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportsGetCount(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iReports = 0
	Local $asNames[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFolder <> "") Then
		$asSplit = StringSplit($sFolder, "/")

		For $i = 1 To $asSplit[0]
			If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
			$oSource = $oSource.getByName($asSplit[$i])
			If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Next
	EndIf

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.ucb.Content") Then ; Report Doc.
			$iReports += 1

		ElseIf $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			ReDim $avFolders[1][2]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.ucb.Content") Then ; Report Doc.
						$iReports += 1

					ElseIf $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						ReDim $avFolders[$iFolders + 1][2]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][2]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, $iReports)
EndFunc   ;==>_LOBase_ReportsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportsGetNames
; Description ...: Retrieve an Array of Report Names contained in a Document.
; Syntax ........: _LOBase_ReportsGetNames(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, retrieves all Report names, including those in sub-folders.
;                  $sFolder             - [optional] a string value. Default is "". The Sub-Folder to return the array of Report names from. See remarks.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder called in $sFolder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Report names contained in this Document. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only an array of main level Report names (not located in folders), or if $bExhaustive is set to True, it will return an array of all Reports contained in the document.
;                  You can narrow the Report name return down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get an array of Reports contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
;                  All Reports located in folders will have the folder path prefixed to the Report name, separated by forward slashes (/). e.g. Folder1/Folder2/Folder3/ReportXYZ.
;                  Calling $bExhaustive with True when searching inside a Folder, will get all Report names from inside that folder, and all sub-folders.
;                  The order of the Report names inside the folders may not necessarily be in proper order, i.e. if there are two sub folders, and Folders inside the first sub-folder, the Reports inside the two folders will be listed first, then the Reports inside the folders inside the first sub-folder.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportsGetNames(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iReports = 0
	Local $asNames[0], $asReports[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFolder <> "") Then
		$asSplit = StringSplit($sFolder, "/")

		For $i = 1 To $asSplit[0]
			If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
			$oSource = $oSource.getByName($asSplit[$i])
			If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Next
		$sFolder &= "/"
	EndIf

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.ucb.Content") Then ; Form Doc.
			If (UBound($asReports) >= $iReports) Then ReDim $asReports[$iReports + 1]
			$asReports[$iReports] = $sFolder & $asNames[$i]
			$iReports += 1

		ElseIf $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			ReDim $avFolders[1][3]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
			$avFolders[0][$iPrefix] = $asNames[$i] & "/"
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.ucb.Content") Then ; Form Doc.
						If (UBound($asReports) >= $iReports) Then ReDim $asReports[$iReports + 1]
						$asReports[$iReports] = $sFolder & $avFolders[$iCount][$iPrefix] & $asFolderList[$k]
						$iReports += 1

					ElseIf $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						ReDim $avFolders[$iFolders + 1][3]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj
						$avFolders[$iFolders][$iPrefix] = $avFolders[$iCount][$iPrefix] & $asFolderList[$k] & "/"

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][3]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($asReports), $asReports)
EndFunc   ;==>_LOBase_ReportsGetNames
