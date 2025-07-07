#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;~ #Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Base
#include "LibreOfficeBase_Internal.au3"

; Other includes for Base

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Closing, Saving, etc. L.O. Base documents.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_DocClose
; _LOBase_DocConnect
; _LOBase_DocCreate
; _LOBase_DocDatabaseType
; _LOBase_DocGetName
; _LOBase_DocGetPath
; _LOBase_DocHasPath
; _LOBase_DocIsActive
; _LOBase_DocIsModified
; _LOBase_DocMaximize
; _LOBase_DocMinimize
; _LOBase_DocOpen
; _LOBase_DocSave
; _LOBase_DocSaveAs
; _LOBase_DocSaveCopy
; _LOBase_DocTableUIClose
; _LOBase_DocTableUIOpenByName
; _LOBase_DocTableUIOpenByObject
; _LOBase_DocTableUIVisible
; _LOBase_DocVisible
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocClose
; Description ...: Close an existing Database Document.
; Syntax ........: _LOBase_DocClose(ByRef $oDoc[, $bSaveChanges = True[, $sSaveName = ""[, $bDeliverOwnership = True]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bSaveChanges        - [optional] a boolean value. Default is True. If true, saves changes if any were made before closing. See remarks.
;                  $sSaveName           - [optional] a string value. Default is "". The file name to save the file as, if the file hasn't been saved before. See Remarks.
;                  $bDeliverOwnership   - [optional] a boolean value. Default is True. If True, deliver ownership of the document Object from the script to LibreOffice, recommended is True.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bSaveChanges not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sSaveName not a String.
;                  @Error 1 @Extended 4 Return 0 = $bDeliverOwnership not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Path Conversion to L.O. URL Failed.
;                  @Error 3 @Extended 2 Return 0 = Error while setting Filter Name properties.
;                  --Document Errors--
;                  @Error 5 @Extended 1 Return 0 = $bSaveChanges set to True, and Document hasn't been assigned a Database type yet. Set it using _LOBase_DocDatabaseType.
;                  @Error 5 @Extended 2 Return 0 = Document hasn't been assigned a Database type yet. Set it using _LOBase_DocDatabaseType.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Success, Document was successfully closed, and was saved to the returned file Path.
;                  @Error 0 @Extended 2 Return String = Success, Document was successfully closed, document's changes were saved to its existing location.
;                  @Error 0 @Extended 3 Return 1 = Success, Document was successfully closed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bSaveChanges is true and the document hasn't been saved yet, the document is saved to the desktop.
;                  You must set the Database type using _LOBase_DocDatabaseType, before you can save the document that hasn't been saved before.
;                  If $sSaveName is undefined, it is saved as an .odb document to the desktop, named Year-Month-Day_Hour-Minute-Second.odb. $sSaveName may be a name only without an extension, in which case the file will be saved in .odb format, you may also include the extension, such as "Test.odb"
; Related .......: _LOBase_DocOpen, _LOBase_DocConnect, _LOBase_DocCreate, _LOBase_DocSaveAs, _LOBase_DocSave, _LOBase_DocDatabaseType
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocClose(ByRef $oDoc, $bSaveChanges = True, $sSaveName = "", $bDeliverOwnership = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__STR_STRIPLEADING = 1, $__STR_STRIPTRAILING = 2
	Local $sDocPath = "", $sSavePath
	Local $aArgs[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bSaveChanges) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSaveName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDeliverOwnership) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If $bSaveChanges And ($oDoc.DataSource.URL() = "") Then Return SetError($__LO_STATUS_DOC_ERROR, 1, 0)

	If ($bSaveChanges = True) Then
		If $oDoc.hasLocation() Then
			$oDoc.store()
			$sDocPath = _LOBase_PathConvert($oDoc.getURL(), $LOB_PATHCONV_PCPATH_RETURN)
			$oDoc.Close($bDeliverOwnership)

			Return SetError($__LO_STATUS_SUCCESS, 2, $sDocPath)

		Else
			If ($oDoc.DataSource.URL() = "") Then Return SetError($__LO_STATUS_DOC_ERROR, 2, 0)

			$sSavePath = @DesktopDir & "\"
			If ($sSaveName = "") Or ($sSaveName = " ") Then
				$sSaveName = @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "-" & @MIN & "-" & @SEC & ".odb"
			EndIf

			$sSaveName = StringStripWS($sSaveName, $__STR_STRIPLEADING + $__STR_STRIPTRAILING)
			If Not StringRegExp($sSaveName, "\Q.odb\E[ ]*$") Then $sSaveName &= ".odb"

			$sSavePath = _LOBase_PathConvert($sSavePath & $sSaveName, 1)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$aArgs[0] = __LOBase_SetPropertyValue("FilterName", "StarOffice XML (Base)")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$oDoc.storeAsURL($sSavePath, $aArgs)
			$oDoc.Close($bDeliverOwnership)

			Return SetError($__LO_STATUS_SUCCESS, 1, _LOBase_PathConvert($sSavePath, $LOB_PATHCONV_PCPATH_RETURN))
		EndIf
	EndIf

	$oDoc.Close($bDeliverOwnership)

	Return SetError($__LO_STATUS_SUCCESS, 3, 1)
EndFunc   ;==>_LOBase_DocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocConnect
; Description ...: Retrieve the Object of an already opened instance of LibreOffice Base.
; Syntax ........: _LOBase_DocConnect($sFile[, $bConnectCurrent = False[, $bConnectAll = False]])
; Parameters ....: $sFile               - a string value. A Full or partial file path, or a full or partial file name. See remarks. Can be an empty string if $bConnectAll or $bConnectCurrent is True.
;                  $bConnectCurrent     - [optional] a boolean value. Default is False. If True, returns the currently active, or last active Document, unless it is not a Calc Document.
;                  $bConnectAll         - [optional] a boolean value. Default is False. If True, returns an array containing all open LibreOffice Base Documents. See remarks.
; Return values .: Success: Object or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFile not a string.
;                  @Error 1 @Extended 2 Return 0 = $bConnectCurrent not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bConnectAll not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating ServiceManager object.
;                  @Error 2 @Extended 2 Return 0 = Error creating Desktop object.
;                  @Error 2 @Extended 3 Return 0 = Error creating enumeration of open documents.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error converting path to Libre Office URL.
;                  --Document Errors--
;                  @Error 5 @Extended 1 Return 0 = No matches found.
;                  @Error 5 @Extended 2 Return 0 = Current Component not a Base Document.
;                  @Error 5 @Extended 3 Return 0 = No open Libre Office documents found.
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Success, The Object for the current, or last active document is returned.
;                  @Error 0 @Extended ? Return Array = Success, An Array of all open LibreOffice Base documents is returned. See remarks. @Extended is set to number of results.
;                  @Error 0 @Extended 3 Return Object = Success, The Object for the document with matching URL is returned.
;                  @Error 0 @Extended 4 Return Object = Success, The Object for the document with matching Title is returned.
;                  @Error 0 @Extended 5 Return Object = Success, A partial Title or Path search found only one match, returning the Object for the found document.
;                  @Error 0 @Extended ? Return Array = Success, An Array of all matching Libre Text documents from a partial Title or Path search. See remarks. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function does not open a connection to the Database, but retrieves an Object for the currently opened Document(s).
;                  $sFile can be either the full Path (Name and extension included; e.g: C:\file\Test.ods Or file:///C:/file/Test.ods) of the document, or the full Title with extension, (e.g: Test.ods), or a partial file path (e.g: file1\file2\Test Or file1\file2 Or file1/file2/ etc.), or a partial name (e.g: test, would match Test1.ods, Test2.xlsx etc.).
;                  Partial file path searches and file name searches, as well as the connect all option, return arrays with three columns per result. ($aArray[0][3]. each result is stored in a separate row;
;                  Row 1, Column 0 contains the Object for that document. e.g. $aArray[0][0] = $oDoc
;                  Row 1, Column 1 contains the Document's full title and extension. e.g. $aArray[0][1] = This Test File.xlsx
;                  Row 1, Column 2 contains the document's full file path. e.g. $aArray[0][2] = C:\Folder1\Folder2\This Test File.xlsx
;                  Row 2, Column 0 contains the Object for the next document. And so on. e.g. $aArray[1][0] = $oDoc2
; Related .......: _LOBase_DocOpen, _LOBase_DocClose, _LOBase_DocCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocConnect($sFile, $bConnectCurrent = False, $bConnectAll = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local Const $STR_STRIPLEADING = 1
	Local $aoConnectAll[1], $aoPartNameSearch[1]
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop
	Local $sServiceName = "com.sun.star.sdb.OfficeDatabaseDocument"

	If Not IsString($sFile) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bConnectCurrent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bConnectAll) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	If Not $oDesktop.getComponents.hasElements() Then Return SetError($__LO_STATUS_DOC_ERROR, 3, 0) ; no L.O open

	$oEnumDoc = $oDesktop.getComponents.createEnumeration()
	If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If $bConnectCurrent Then
		$oDoc = $oDesktop.currentComponent()

		Return ($oDoc.supportsService($sServiceName)) ? (SetError($__LO_STATUS_SUCCESS, 1, $oDoc)) : (SetError($__LO_STATUS_DOC_ERROR, 2, 0))
	EndIf

	If $bConnectAll Then
		ReDim $aoConnectAll[1][3]
		$iCount = 0
		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService($sServiceName) Then
				ReDim $aoConnectAll[$iCount + 1][3]
				$aoConnectAll[$iCount][0] = $oDoc
				$aoConnectAll[$iCount][1] = $oDoc.Title()
				$aoConnectAll[$iCount][2] = _LOBase_PathConvert($oDoc.getURL(), $LOB_PATHCONV_PCPATH_RETURN)
				$iCount += 1
			EndIf
			Sleep(10)
		WEnd

		Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoConnectAll)
	EndIf

	$sFile = StringStripWS($sFile, $STR_STRIPLEADING)
	If StringInStr($sFile, "\") Then $sFile = _LOBase_PathConvert($sFile, $LOB_PATHCONV_OFFICE_RETURN) ; Convert to L.O File path.
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If StringInStr($sFile, "file:///") Then ; URL/Path and Name search

		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()

			If ($oDoc.getURL() == $sFile) Then Return SetError($__LO_STATUS_SUCCESS, 3, $oDoc) ; Match
		WEnd

		Return SetError($__LO_STATUS_DOC_ERROR, 1, 0) ; no match

	Else
		If Not StringInStr($sFile, "/") And StringInStr($sFile, ".") Then ; Name with extension only search
			While $oEnumDoc.hasMoreElements()
				$oDoc = $oEnumDoc.nextElement()
				If StringInStr($oDoc.Title, $sFile) Then Return SetError($__LO_STATUS_SUCCESS, 4, $oDoc) ; Match
			WEnd

			Return SetError($__LO_STATUS_DOC_ERROR, 1, 0) ; no match
		EndIf

		$iCount = 0 ; partial name or partial URL search
		ReDim $aoPartNameSearch[$iCount + 1][3]

		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If StringInStr($sFile, "/") Then
				If StringInStr($oDoc.getURL(), $sFile) Then
					ReDim $aoPartNameSearch[$iCount + 1][3]
					$aoPartNameSearch[$iCount][0] = $oDoc
					$aoPartNameSearch[$iCount][1] = $oDoc.Title
					$aoPartNameSearch[$iCount][2] = _LOBase_PathConvert($oDoc.getURL, $LOB_PATHCONV_PCPATH_RETURN)
					$iCount += 1
				EndIf

			Else
				If StringInStr($oDoc.Title, $sFile) Then
					ReDim $aoPartNameSearch[$iCount + 1][3]
					$aoPartNameSearch[$iCount][0] = $oDoc
					$aoPartNameSearch[$iCount][1] = $oDoc.Title
					$aoPartNameSearch[$iCount][2] = _LOBase_PathConvert($oDoc.getURL, $LOB_PATHCONV_PCPATH_RETURN)
					$iCount += 1
				EndIf
			EndIf
		WEnd
		If IsString($aoPartNameSearch[0][1]) Then
			If (UBound($aoPartNameSearch) = 1) Then

				Return SetError($__LO_STATUS_SUCCESS, 5, $aoPartNameSearch[0][0]) ; matches

			Else

				Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoPartNameSearch) ; matches
			EndIf

		Else

			Return SetError($__LO_STATUS_DOC_ERROR, 1, 0) ; no match
		EndIf
	EndIf
EndFunc   ;==>_LOBase_DocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocCreate
; Description ...: Open a new Libre Office Base Document.
; Syntax ........: _LOBase_DocCreate([$bForceNew = True[, $bHidden = False[, $bWizard = False]]])
; Parameters ....: $bForceNew           - [optional] a boolean value. Default is True. If True, force opening a new Base Document instead of checking for a usable blank.
;                  $bHidden             - [optional] a boolean value. Default is False. If True opens the new document invisible or changes the existing document to invisible.
;                  $bWizard             - [optional] a boolean value. Default is False. If True, opens the Create a Database Document wizard. See remarks.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bForceNew not a Boolean.
;                  @Error 1 @Extended 2 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bWizard not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bWizar and $bHidden both set to True.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failure Creating Object com.sun.star.ServiceManager.
;                  @Error 2 @Extended 2 Return 0 = Failure Creating Object com.sun.star.frame.Desktop.
;                  @Error 2 @Extended 3 Return 0 = Failed to enumerate available documents.
;                  @Error 2 @Extended 4 Return 0 = Failure Creating New Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Document Object is still returned. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHidden
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Successfully connected to an existing Document. Returning Document's Object
;                  @Error 0 @Extended 2 Return Object = Successfully created a new document. Returning Document's Object
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bWizard is True, $bHidden must be false.
;                  If $bWizard is True, the function will not return until the user either cancels or completes the wizard. If the user cancels, an error will result.
;                  You must set the Database type using _LOBase_DocDatabaseType, before you can save the document.
; Related .......: LOBase_DocOpen, LOBase_DocClose, LOBase_DocConnect, _LOBase_DocDatabaseType
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocCreate($bForceNew = True, $bHidden = False, $bWizard = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $aArgs[1]
	Local $iError = 0
	Local $oServiceManager, $oDesktop, $oDoc, $oEnumDoc
	Local $sServiceName = "com.sun.star.sdb.DatabaseDocument"

	If Not IsBool($bForceNew) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bWizard) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $bWizard And $bHidden Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$aArgs[0] = __LOBase_SetPropertyValue("Hidden", $bHidden)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; If not force new, and L.O. pages exist then see if there are any blank Base documents to use.
	If Not $bForceNew And $oDesktop.getComponents.hasElements() Then
		$oEnumDoc = $oDesktop.getComponents.createEnumeration()
		If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService($sServiceName) _
					And Not ($oDoc.hasLocation() And $oDoc.isReadOnly()) And ($oDoc.WordCount() = 0) Then
				$oDoc.CurrentController.Frame.ContainerWindow.Visible = ($bHidden) ? (False) : (True) ; opposite value of $bHidden.
				$iError = ($oDoc.CurrentController.Frame.isHidden() = $bHidden) ? ($iError) : (BitOR($iError, 1))

				Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))
			EndIf
		WEnd
	EndIf

	If Not IsObj($aArgs[0]) Then $iError = BitOR($iError, 1)

	If $bWizard Then
		$oDoc = $oDesktop.loadComponentFromURL("private:factory/sdatabase?Interactive", "_blank", $iURLFrameCreate, $aArgs)

	Else
		$oDoc = $oDesktop.loadComponentFromURL("private:factory/sdatabase", "_blank", $iURLFrameCreate, $aArgs)
	EndIf

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOBase_DocCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocDatabaseType
; Description ...: Set or Retrieve a Base Document's Database Type.
; Syntax ........: _LOBase_DocDatabaseType(ByRef $oDoc[, $sType = "sdbc:embedded:hsqldb"[, $bOverwrite = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sType               - [optional] a string value. Default is Default. Default is "sdbc:embedded:hsqldb". The Database Type string to set the document to. See remarks.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, an existing Database type will be overwritten. See remarks.
; Return values .: Success: 1 or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sType not a String.
;                  @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Database type.
;                  @Error 3 @Extended 2 Return 0 = $bOverwrite is set to False, and Document's Database type is already set.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sType
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. $sType set to Null, returning current Database type as a String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: I have not investigated the various settings for each Database type therefore I have no checks for right or wrong values, if you know the appropriate string to use you can set $sType to the desired setting, but make sure you know what you are doing. "sdbc:embedded:hsqldb" is the default setting for LibreOffice, which creates an embedded HSQLDB Base Document. The Type format is as follows jdbc:subprotocol:subname or sdbc:subprotocol:subname.
;                  I am not knowledgeable enough to know if changing Database types works, or if it is advisable, therefore I made the setting $bOverwrite. If $bOverwrite is False it prevents the user from setting the Database type if one is already set for the document.
;                  Call this function with $sType set to Null to get the current Database type.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocDatabaseType(ByRef $oDoc, $sType = Default, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sDataType

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sDataType = $oDoc.DataSource.URL()
	If Not IsString($sDataType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sType = Null) Then

		Return SetError($__LO_STATUS_SUCCESS, 1, $sDataType)

	ElseIf ($sType = Default) Then
		$sType = "sdbc:embedded:hsqldb"
	EndIf

	If Not IsString($sType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sDataType <> "jdbc:") And ($bOverwrite = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.DataSource.URL = $sType
	If ($oDoc.DataSource.URL() <> $sType) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DocDatabaseType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocGetName
; Description ...: Retrieve the document's name.
; Syntax ........: _LOBase_DocGetName(ByRef $oDoc[, $bReturnFull = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bReturnFull         - [optional] a boolean value. Default is False. If True, the full window title is returned, such as is used by AutoIt window related functions.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bReturnFull not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returns the document's current Name/Title
;                  @Error 0 @Extended 1 Return String = Success. Returns the document's current Window Title, which includes the document name and usually: "-LibreOffice Base".
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocGetName(ByRef $oDoc, $bReturnFull = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnFull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sName = ($bReturnFull = True) ? ($oDoc.CurrentController.Frame.Title()) : ($oDoc.Title())

	Return ($bReturnFull = True) ? (SetError($__LO_STATUS_SUCCESS, 1, $sName)) : (SetError($__LO_STATUS_SUCCESS, 0, $sName))
EndFunc   ;==>_LOBase_DocGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocGetPath
; Description ...: Returns a Document's current save path.
; Syntax ........: _LOBase_DocGetPath(ByRef $oDoc[, $bReturnLibreURL = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bReturnLibreURL     - [optional] a boolean value. Default is False. If True, returns a path in Libre Office URL format, else false returns a regular Windows path.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bReturnLibreURL not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = Document has no save path.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error converting Libre URL to Computer path format.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returns the P.C. path to the current document's save path.
;                  @Error 0 @Extended 1 Return String = Success. Returns the Libre Office URL to the current document's save path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_PathConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocGetPath(ByRef $oDoc, $bReturnLibreURL = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sPath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnLibreURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.hasLocation() Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($bReturnLibreURL = True) Then
		$sPath = $oDoc.URL()

	Else
		$sPath = $oDoc.URL()
		$sPath = _LOBase_PathConvert($sPath, $LOB_PATHCONV_PCPATH_RETURN)
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return ($bReturnLibreURL = True) ? (SetError($__LO_STATUS_SUCCESS, 1, $sPath)) : (SetError($__LO_STATUS_SUCCESS, 0, $sPath))
EndFunc   ;==>_LOBase_DocGetPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocHasPath
; Description ...: Returns whether a document has been saved to a location already or not.
; Syntax ........: _LOBase_DocHasPath(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the document has a save location. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocHasPath(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.hasLocation())
EndFunc   ;==>_LOBase_DocHasPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocIsActive
; Description ...: Tests if called document is the active document of other Libre windows.
; Syntax ........: _LOBase_DocIsActive(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if document is the currently active Libre window. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This does NOT test if the document is the current active window in Windows, it only tests if the document is the current active document among other Libre Office documents.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocIsActive(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.CurrentController.Frame.isActive())
EndFunc   ;==>_LOBase_DocIsActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocIsModified
; Description ...: Test whether the document has been modified since being created or since the last save.
; Syntax ........: _LOBase_DocIsModified(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the document has been modified since last being saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocIsModified(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.isModified())
EndFunc   ;==>_LOBase_DocIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocMaximize
; Description ...: Maximize or restore a document.
; Syntax ........: _LOBase_DocMaximize(ByRef $oDoc[, $bMaximize = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bMaximize           - [optional] a boolean value. Default is Null. If True, document window is maximized, else if false, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bMaximize not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully maximized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMaximize set to Null, returning boolean indicating if Document is currently maximized (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bMaximize is set to Null, returns a Boolean indicating if document is currently maximized (True).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocMaximize(ByRef $oDoc, $bMaximize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bMaximize = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.IsMaximized())

	If Not IsBool($bMaximize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.IsMaximized = $bMaximize

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DocMaximize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocMinimize
; Description ...: Minimize or restore a document.
; Syntax ........: _LOBase_DocMinimize(ByRef $oDoc[, $bMinimize = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bMinimize           - [optional] a boolean value. Default is Null. If True, document window is minimized, else if false, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bMinimize not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully minimized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMinimize set to Null, returning boolean indicating if Document is currently minimized (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bMinimize is set to Null, returns a Boolean indicating if document is currently minimized (True).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocMinimize(ByRef $oDoc, $bMinimize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bMinimize = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.IsMinimized())

	If Not IsBool($bMinimize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.IsMinimized = $bMinimize

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DocMinimize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocOpen
; Description ...: Open an existing Database Document.
; Syntax ........: _LOBase_DocOpen($sFilePath[, $bConnectIfOpen = True[, $bHidden = Null[, $bReadOnly = Null[, $sPassword = Null[, $bLoadAsTemplate = Null[, $sFilterName = Null]]]]]])
; Parameters ....: $sFilePath           - a string value. Full path and filename of the file to be opened.
;                  $bConnectIfOpen      - [optional] a boolean value. Default is True(Connect). Whether to connect to the requested document if it is already open. See remarks.
;                  $bHidden             - [optional] a boolean value. Default is Null. If true, opens the document invisibly.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If true, opens the document as read-only.
;                  $sPassword           - [optional] a string value. Default is Null. The password that was used to read-protect the document, if any.
;                  $bLoadAsTemplate     - [optional] a boolean value. Default is Null. If true, opens the document as a Template, i.e. an untitled copy of the specified document is made instead of modifying the original document.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFilePath not string, or file not found.
;                  @Error 1 @Extended 2 Return 0 = Error converting file path to URL path.
;                  @Error 1 @Extended 3 Return 0 = $bConnectIfOpen not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $sPassword not a string.
;                  @Error 1 @Extended 7 Return 0 = $bLoadAsTemplate not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create ServiceManager Object
;                  @Error 2 @Extended 2 Return 0 = Failed to create Desktop Object
;                  @Error 2 @Extended 3 Return 0 = Failed opening or connecting to document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHidden
;                  |                               2 = Error setting $bReadOnly
;                  |                               4 = Error setting $sPassword
;                  |                               8 = Error setting $bLoadAsTemplate
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Successfully connected to requested Document without requested parameters. Returning Document's Object.
;                  @Error 0 @Extended 2 Return Object = Successfully opened requested Document with requested parameters. Returning Document's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Any parameters (Hidden, template etc.,) will not be applied when connecting to a document.
; Related .......: _LOBase_DocCreate, _LOBase_DocClose, _LOBase_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocOpen($sFilePath, $bConnectIfOpen = True, $bHidden = Null, $bReadOnly = Null, $sPassword = Null, $bLoadAsTemplate = Null)
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $iError = 0
	Local $oDoc, $oServiceManager, $oDesktop
	Local $aoProperties[0]
	Local $vProperty
	Local $sFileURL

	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsString($sFilePath) Or Not FileExists($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sFileURL = _LOBase_PathConvert($sFilePath, $LOB_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bConnectIfOpen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not __LOBase_VarsAreNull($bHidden, $bReadOnly, $sPassword, $bLoadAsTemplate) Then
		If ($bHidden <> Null) Then
			If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$vProperty = __LOBase_SetPropertyValue("Hidden", $bHidden)
			If @error Then $iError = BitOR($iError, 1)
			If Not BitAND($iError, 1) Then __LOBase_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($bReadOnly <> Null) Then
			If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$vProperty = __LOBase_SetPropertyValue("ReadOnly", $bReadOnly)
			If @error Then $iError = BitOR($iError, 2)
			If Not BitAND($iError, 2) Then __LOBase_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($sPassword <> Null) Then
			If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$vProperty = __LOBase_SetPropertyValue("Password", $sPassword)
			If @error Then $iError = BitOR($iError, 4)
			If Not BitAND($iError, 4) Then __LOBase_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($bLoadAsTemplate <> Null) Then
			If Not IsBool($bLoadAsTemplate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$vProperty = __LOBase_SetPropertyValue("AsTemplate", $bLoadAsTemplate)
			If @error Then $iError = BitOR($iError, 8)
			If Not BitAND($iError, 8) Then __LOBase_AddTo1DArray($aoProperties, $vProperty)
		EndIf
	EndIf

	If $bConnectIfOpen Then $oDoc = _LOBase_DocConnect($sFilePath)
	If IsObj($oDoc) Then Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))

	$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOBase_DocOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocSave
; Description ...: Save any changes made to a Document.
; Syntax ........: _LOBase_DocSave(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Document is Read Only or Document has no save location, try SaveAs.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Document Successfully saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You must set the DataBase type using _LOBase_DocDatabaseType, before you can save the document.
; Related .......: _LOBase_DocSaveAs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocSave(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDoc.hasLocation Or $oDoc.isReadOnly Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oDoc.store()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DocSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocSaveAs
; Description ...: Save a Document with the specified file name to the path specified with any parameters called.
; Syntax ........: _LOBase_DocSaveAs(ByRef $oDoc, $sFilePath[, $bOverwrite = Null[, $sPassword = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sFilePath           - a string value. Full path to save the document to, including Filename and extension.
;                  $bOverwrite          - [optional] a boolean value. Default is Null. If True, the existing file will be overwritten.
;                  $sPassword           - [optional] a string value. Default is Null. Sets a password for the document. (Not all file formats can have a Password set). Null or "" (blank string) = No Password.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFilePath not a String.
;                  @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $sPassword not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Converting Path to/from L.O. URL
;                  @Error 3 @Extended 2 Return 0 = Error setting FilterName Property
;                  @Error 3 @Extended 3 Return 0 = Error setting Overwrite Property
;                  @Error 3 @Extended 4 Return 0 = Error setting Password Property
;                  --Document Errors--
;                  @Error 5 @Extended 1 Return 0 = Document hasn't been assigned a Database type yet. Set it using _LOBase_DocDatabaseType.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Successfully Saved the document. Returning document save path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Alters original save path (if there was one) to the new path.
;                  If ".odb" extension is not present, it is appended to the save path.
;                  You must set the Database type using _LOBase_DocDatabaseType, before you can save the document.
; Related .......: _LOBase_DocSave, _LOBase_DocSaveCopy, _LOBase_DocDatabaseType
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocSaveAs(ByRef $oDoc, $sFilePath, $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__STR_STRIPLEADING = 1, $__STR_STRIPTRAILING = 2
	Local $aProperties[1]
	Local $sSavePath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oDoc.DataSource.URL() = "") Then Return SetError($__LO_STATUS_DOC_ERROR, 1, 0)

	$sFilePath = StringStripWS($sFilePath, $__STR_STRIPLEADING + $__STR_STRIPTRAILING)
	If Not StringRegExp($sFilePath, "\Q.odb\E[ ]*$") Then $sFilePath &= ".odb"

	$sFilePath = _LOBase_PathConvert($sFilePath, $LOB_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aProperties[0] = __LOBase_SetPropertyValue("FilterName", "StarOffice XML (Base)")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOBase_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	EndIf

	If $sPassword <> Null Then
		If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOBase_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oDoc.storeAsURL($sFilePath, $aProperties)

	$sSavePath = _LOBase_PathConvert($sFilePath, $LOB_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOBase_DocSaveAs

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocSaveCopy
; Description ...: Save a copy of a Document to the path and file name specified, without modifying the original save location.
; Syntax ........: _LOBase_DocSaveCopy(ByRef $oDoc, $sFilePath[, $bOverwrite = Null[, $sPassword = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sFilePath           - a string value. Full path to save the document to, including Filename and extension. See Remarks.
;                  $bOverwrite          - [optional] a boolean value. Default is Null. If True, file will be overwritten.
;                  $sPassword           - [optional] a string value. Default is Null. Password String to set for the document. (Not all file formats can have a Password set). "" (blank string) or Null = No Password.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFilePath not a String.
;                  @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $sPassword not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Converting Path to/from L.O. URL
;                  @Error 3 @Extended 2 Return 0 = Error setting FilterName Property
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Error setting Overwrite Property
;                  @Error 4 @Extended 2 Return 0 = Error setting Password Property
;                  --Document Errors--
;                  @Error 5 @Extended 1 Return 0 = $bSaveChanges set to True, and Document hasn't been assigned a Database type yet. Set it using _LOBase_DocDatabaseType.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning save path for exported document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Does not alter the original save path (if there was one), saves a copy of the document to the new path.
; Related .......: _LOBase_DocSave, _LOBase_DocSaveAs, _LOBase_DocDatabaseType
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocSaveCopy(ByRef $oDoc, $sFilePath, $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aProperties[3]
	Local $sSavePath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oDoc.DataSource.URL() = "") Then Return SetError($__LO_STATUS_DOC_ERROR, 1, 0)

	$sFilePath = _LOBase_PathConvert($sFilePath, $LOB_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aProperties[0] = __LOBase_SetPropertyValue("FilterName", "StarOffice XML (Base)")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOBase_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)
	EndIf

	If ($sPassword <> Null) Then
		If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOBase_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0)
	EndIf

	$oDoc.storeToURL($sFilePath, $aProperties)

	$sSavePath = _LOBase_PathConvert($sFilePath, $LOB_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOBase_DocSaveCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocTableUIClose
; Description ...: Close a Table User Interface window.
; Syntax ........: _LOBase_DocTableUIClose(ByRef $oTableUI[, $bDeliverOwnership = True])
; Parameters ....: $oTableUI            - [in/out] an object. A Table User Interface Object from a previous _LOBase_DocTableUIOpenByName or _LOBase_DocTableUIOpenByObject function.
;                  $bDeliverOwnership   - [optional] a boolean value. Default is True. If True, deliver ownership of the Table UI Object from the script to LibreOffice, recommended is True.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTableUI not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bDeliverOwnership not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully closed the Table User Interface window.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocTableUIClose(ByRef $oTableUI, $bDeliverOwnership = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oTableUI) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bDeliverOwnership) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTableUI.Frame.close($bDeliverOwnership)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DocTableUIClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocTableUIOpenByName
; Description ...: Open a Table's User Interface windows either in design mode or viewing mode.
; Syntax ........: _LOBase_DocTableUIOpenByName(ByRef $oDoc, ByRef $oConnection, $sTable[, $bEdit = False[, $bVisible = True]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sTable              - a string value. The Table's name.
;                  $bEdit               - [optional] a boolean value. Default is False. If True, the Table is opened in editing mode to add or remove columns. If False, the table is opened in data viewing mode, to modify Table Data.
;                  $bVisible            - [optional] a boolean value. Default is True. If True, the UI window will be visible.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 3 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 4 Return 0 = $sTable not a String.
;                  @Error 1 @Extended 5 Return 0 = $bEdit not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = No Table with name called in $sTable found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Tables Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to create a Connection to Database.
;                  @Error 3 @Extended 4 Return 0 = Failed to open Table UI.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully opened Table's User Interface, returning its object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DocTableUIOpenByObject, _LOBase_DocTableUIClose, _LOBase_DocTableUIVisible
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocTableUIOpenByName(ByRef $oDoc, ByRef $oConnection, $sTable, $bEdit = False, $bVisible = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTables, $oTableUI
	Local Const $__LOB_OBJ_TYPE_TABLE = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bEdit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oTables = $oConnection.getTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oTables.hasByName($sTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If Not $oDoc.CurrentController.isConnected() Then $oDoc.CurrentController.connect()
	If Not $oDoc.CurrentController.isConnected() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oTableUI = $oDoc.CurrentController.loadComponent($__LOB_OBJ_TYPE_TABLE, $sTable, $bEdit)

	If Not IsObj($oTableUI) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$oTableUI.Frame.ContainerWindow.Visible = $bVisible

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTableUI)
EndFunc   ;==>_LOBase_DocTableUIOpenByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocTableUIOpenByObject
; Description ...: Open a Table's User Interface windows either in design mode or viewing mode.
; Syntax ........: _LOBase_DocTableUIOpenByObject(ByRef $oDoc, ByRef $oConnection, ByRef $oTable[, $bEdit = False[, $bVisible = True]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
;                  $bEdit               - [optional] a boolean value. Default is False. If True, the Table is opened in editing mode to add or remove columns. If False, the table is opened in data viewing mode, to modify Table Data.
;                  $bVisible            - [optional] a boolean value. Default is True. If True, the UI window will be visible.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 3 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 4 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 5 Return 0 = $bEdit not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bVisible not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Table Name.
;                  @Error 3 @Extended 3 Return 0 = Failed to create a Connection to Database.
;                  @Error 3 @Extended 4 Return 0 = Failed to open Table UI.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully opened Table's User Interface, returning its object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DocTableUIOpenByName, _LOBase_DocTableUIClose, _LOBase_DocTableUIVisible
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocTableUIOpenByObject(ByRef $oDoc, ByRef $oConnection, ByRef $oTable, $bEdit = False, $bVisible = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTableUI
	Local $sTable
	Local Const $__LOB_OBJ_TYPE_TABLE = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bEdit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sTable = $oTable.Name()
	If Not IsString($sTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not $oDoc.CurrentController.isConnected() Then $oDoc.CurrentController.connect()
	If Not $oDoc.CurrentController.isConnected() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oTableUI = $oDoc.CurrentController.loadComponent($__LOB_OBJ_TYPE_TABLE, $sTable, $bEdit)

	If Not IsObj($oTableUI) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$oTableUI.Frame.ContainerWindow.Visible = $bVisible

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTableUI)
EndFunc   ;==>_LOBase_DocTableUIOpenByObject

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocTableUIVisible
; Description ...: Set or Retrieve Table UI Visibility.
; Syntax ........: _LOBase_DocTableUIVisible(ByRef $oTableUI[, $bVisible = Null])
; Parameters ....: $oTableUI            - [in/out] an object. A Table User Interface Object from a previous _LOBase_DocTableUIOpenByName or _LOBase_DocTableUIOpenByObject function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the Table UI Window is visible.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTableUI not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current visibility setting.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. All optional parameters were set to Null, returning current visibility setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocTableUIVisible(ByRef $oTableUI, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oTableUI) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bVisible = Null) Then
		$bVisible = $oTableUI.Frame.ContainerWindow.IsVisible()
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $bVisible)
	EndIf

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTableUI.Frame.ContainerWindow.Visible = $bVisible
	If Not ($oTableUI.Frame.ContainerWindow.IsVisible() = $bVisible) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DocTableUIVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocVisible
; Description ...: Set or retrieve the current visibility of a document.
; Syntax ........: _LOBase_DocVisible(ByRef $oDoc[, $bVisible = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the document is visible.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Error setting $bVisible.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. $bVisible successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. Returning current visibility state of the Document, True if visible, false if invisible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call $bVisible with Null to return the current visibility setting.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocVisible(ByRef $oDoc, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bVisible = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.isVisible())

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.Visible = $bVisible
	$iError = ($oDoc.CurrentController.Frame.ContainerWindow.isVisible() = $bVisible) ? (0) : (1)

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOBase_DocVisible
