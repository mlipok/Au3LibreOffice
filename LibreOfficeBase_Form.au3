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
; Description ...: Provides basic functionality through AutoIt for Adding, Deleting, and modifying, etc. L.O. Base Forms.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; Notes .........: Forms are simply Writer Documents stored internally in an obd file. Most _LOWriter_* functions should work with a form document object also.
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_FormClose
; _LOBase_FormConnect
; _LOBase_FormCopy
; _LOBase_FormCreate
; _LOBase_FormDelete
; _LOBase_FormDocVisible
; _LOBase_FormExists
; _LOBase_FormFolderCopy
; _LOBase_FormFolderCreate
; _LOBase_FormFolderDelete
; _LOBase_FormFolderExists
; _LOBase_FormFolderRename
; _LOBase_FormFoldersGetCount
; _LOBase_FormFoldersGetNames
; _LOBase_FormIsModified
; _LOBase_FormOpen
; _LOBase_FormRename
; _LOBase_FormSave
; _LOBase_FormsGetCount
; _LOBase_FormsGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormClose
; Description ...: Close an opened Form Document.
; Syntax ........: _LOBase_FormClose(ByRef $oFormDoc[, $bForceClose = False])
; Parameters ....: $oFormDoc            - [in/out] an object. A Form Document object returned by a previous _LOBase_FormOpen, _LOBase_FormConnect, or _LOBase_FormCreate function.
;                  $bForceClose         - [optional] a boolean value. Default is False. If True, the Form document will be closed regardless if there are unsaved changes. See remarks.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bForceClose not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = Document called in $oFormDoc has not been saved to a Base Document yet.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Document has been modified and not saved, and $bForceClose is False.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify Form Document's Unique Identifier String.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify Form Document's name.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Form and Folder names.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Form or Folder Object.
;                  @Error 3 @Extended 7 Return 0 = Failed to retrieve Array of Form and Folder names for Sub-Folder.
;                  @Error 3 @Extended 8 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  @Error 3 @Extended 9 Return 0 = Failed to identify Form in Parent Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning a Boolean value of whether the Form Document was successfully closed (True), or not.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If there are unsaved changes in the document when close is called, and $bForceClose is True, they will be lost.
; Related .......: _LOBase_FormOpen, _LOBase_FormConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormClose(ByRef $oFormDoc, $bForceClose = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn
	Local $oForm, $oSource, $oObj
	Local $sUniqueString, $sTitle
	Local $iCount = 0, $iFolders = 1
	Local $asResult[0], $asNames[0], $asFolderList[0]
	Local $avFolders[0][2]
	Local Enum $iName, $iObj, $iPrefix
	Local Const $__STR_REGEXPARRAYMATCH = 1

	If Not IsObj($oFormDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bForceClose) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oFormDoc.isModified() And Not $bForceClose Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If Not $oFormDoc.hasLocation() Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	; Retrieve the unique name of the form.
	; file:///C:/Folder1/Autoit/Libre%20Office%20Base%20Info/Testing.odb/Obj61/
	$asResult = StringRegExp($oFormDoc.NameSpace(), "/([a-zA-Z0-9]+)/$", $__STR_REGEXPARRAYMATCH)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$sUniqueString = $asResult[0]

	; "Testing.odb : abc1Form2"
	$asResult = StringRegExp($oFormDoc.Title(), "\: (.+)$", $__STR_REGEXPARRAYMATCH)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$sTitle = $asResult[0]

	$oSource = $oFormDoc.Parent.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	If $oSource.hasByName($sTitle) And $oSource.getByName($sTitle).supportsService("com.sun.star.ucb.Content") And ($oSource.getByName($sTitle).PersistentName() = $sUniqueString) Then
		$oForm = $oSource.getByName($sTitle)

	Else

		For $i = 0 To UBound($asNames) - 1
			$oObj = $oSource.getByName($asNames[$i])
			If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

			If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
				ReDim $avFolders[1][2]
				$avFolders[0][$iName] = $asNames[$i]
				$avFolders[0][$iObj] = $oObj
			EndIf

			While ($iCount < UBound($avFolders))
				If $avFolders[$iCount][$iObj].hasByName($sTitle) And $avFolders[$iCount][$iObj].getByName($sTitle).supportsService("com.sun.star.ucb.Content") And _
						($avFolders[$iCount][$iObj].getByName($sTitle).PersistentName() = $sUniqueString) Then
					$oForm = $avFolders[$iCount][$iObj].getByName($sTitle)
					ExitLoop
				EndIf

				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

					If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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

	If Not IsObj($oForm) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 9, 0)

	If $oFormDoc.isModified() Then $oFormDoc.Modified = False ; Set modified to false, so the user wont be prompted.

	$bReturn = $oForm.Close()

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_FormClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormConnect
; Description ...: Retrieve an Object for the currently open Form or Forms.
; Syntax ........: _LOBase_FormConnect([$bConnectCurrent = True])
; Parameters ....: $bConnectCurrent     - [optional] a boolean value. Default is True. If True, Returns an Object for the last active Form. Else an array of all Open Forms. See Remarks.
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
;                  @Error 3 @Extended 2 Return 0 = Current LibreOffice window is not a Form Document.
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Success. Connected to the currently active window, returning the Form Document Object.
;                  @Error 0 @Extended 2 Return Array = Success. Returning a Two columned Array with all open Form Documents. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Returned array when connecting to all open Form Documents returns an array with Two columns per result. ($aArray[0][2]). Each result is stored in a separate row;
;                  Row 1, Column 0 contain the Object for that document. e.g. $aArray[0][0] = $oDoc
;                  Row 1, Column 1 contains the Document's full title with extension and the Form Name, separated by a colon. e.g. $aArray[0][1] = "Testing.odb : Form1"
;                  Row 2, Column 0 contain the Object for the next document. And so on. e.g. $aArray[1][0] = $oDoc2
; Related .......: _LOBase_FormOpen, _LOBase_FormClose
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormConnect($bConnectCurrent = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local $aoConnectAll[0][2]
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop
	Local $sServiceName = "com.sun.star.text.TextDocument"

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

		Return (($oDoc.supportsService($sServiceName) And IsObj($oDoc.Parent()))) ? (SetError($__LO_STATUS_SUCCESS, 1, $oDoc)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0))

	Else

		ReDim $aoConnectAll[1][2]
		$iCount = 0
		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If ($oDoc.supportsService($sServiceName) And IsObj($oDoc.Parent())) Then ; If Parent is present, it should be a Database Form.

				ReDim $aoConnectAll[$iCount + 1][2]
				$aoConnectAll[$iCount][0] = $oDoc
				$aoConnectAll[$iCount][1] = $oDoc.Title()
				$iCount += 1
			EndIf
			Sleep(10)
		WEnd

		Return SetError($__LO_STATUS_SUCCESS, 2, $aoConnectAll)
	EndIf
EndFunc   ;==>_LOBase_FormConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormCopy
; Description ...: Create a copy of an existing Form.
; Syntax ........: _LOBase_FormCopy(ByRef $oDoc, ByRef $oConnection, $sInputForm, $sOutputForm)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sInputForm          - a string value. The Name of the Form to Copy. Also the Sub-directory the Form is in. See Remarks.
;                  $sOutputForm       - a string value. The Name of the Form to Create. Also the Sub-directory to place the Form in. See Remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sInputForm not a String.
;                  @Error 1 @Extended 4 Return 0 = $sOutputForm not a String.
;                  @Error 1 @Extended 5 Return 0 = Folder name called in $sInputForm not found.
;                  @Error 1 @Extended 6 Return 0 = Requested Form not found.
;                  @Error 1 @Extended 7 Return 0 = Form name called in $sInputForm not a Form.
;                  @Error 1 @Extended 8 Return 0 = Folder name called in $sOutputForm not found.
;                  @Error 1 @Extended 9 Return 0 = Form already exists with called name in Destination.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.sdb.DocumentDefinition" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Source Folder Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Form Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 6 Return 0 = Failed to insert copied Form.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Copied Form successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To copy a Form located inside a folder, the Form name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to create FormXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sInputForm with the following path: Folder1/Folder2/Folder3/FormXYZ.
;                  To create a Form inside a folder, the Form name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to create FormXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sOutputForm with the following path: Folder1/Folder2/Folder3/FormXYZ.
;                  If only a name is called in $sOutputForm, the Form will be created in the main directory, i.e. not inside of any folders.
; Related .......: _LOBase_FormDelete, _LOBase_FormCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormCopy(ByRef $oDoc, ByRef $oConnection, $sInputForm, $sOutputForm)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oDestination, $oFormDef, $oDocDef
	Local $asSplit[0]
	Local $aArgs[3]
	Local $sSourceForm, $sDestForm

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sInputForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sOutputForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSource = $oDoc.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$asSplit = StringSplit($sInputForm, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	Next

	$sSourceForm = $asSplit[$asSplit[0]] ; Last element of Array will be the Input Form's name to Copy

	If Not $oSource.hasByName($sSourceForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oFormDef = $oSource.getByName($sSourceForm)
	If Not IsObj($oFormDef) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	If Not $oFormDef.supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$oDestination = $oDoc.FormDocuments()
	If Not IsObj($oDestination) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$asSplit = StringSplit($sOutputForm, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oDestination.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$oDestination = $oDestination.getByName($asSplit[$i])
		If Not IsObj($oDestination) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	Next

	$sDestForm = $asSplit[$asSplit[0]] ; Last element of Array will be the Output Form name to Create

	If $oDestination.hasByName($sDestForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	$aArgs[0] = __LOBase_SetPropertyValue("Name", $sDestForm)
	$aArgs[1] = __LOBase_SetPropertyValue("ActiveConnection", $oConnection)
	$aArgs[2] = __LOBase_SetPropertyValue("EmbeddedObject", $oFormDef)

	$oDocDef = $oDestination.createInstanceWithArguments("com.sun.star.sdb.DocumentDefinition", $aArgs)
	If Not IsObj($oDocDef) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDestination.insertbyName($sDestForm, $oDocDef)
	If Not $oDestination.hasByName($sDestForm) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FormCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormCreate
; Description ...: Create and Insert a new Form Document into a Base Document.
; Syntax ........: _LOBase_FormCreate(ByRef $oDoc, ByRef $oConnection, $sForm[, $bOpen = False[, $bDesign = True[, $bVisible = True]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sForm               - a string value. The Name of the Form to Create. Also the Sub-directory to place the form in. See Remarks.
;                  $bOpen               - [optional] a boolean value. Default is False. If True, the new Form will be opened.
;                  $bDesign             - [optional] a boolean value. Default is True. If True, and $bOpen is True, the Form will be opened in Design mode. Else in Form mode.
;                  $bVisible            - [optional] a boolean value. Default is True. If True, the Form will be visible.
; Return values .: Success: 1 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sForm not a String.
;                  @Error 1 @Extended 4 Return 0 = $bOpen not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bDesign not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = Folder or Sub-Folder not found.
;                  @Error 1 @Extended 8 Return 0 = Name called in $sForm already exists in Folder.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.ServiceManager Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create com.sun.star.frame.Desktop Object.
;                  @Error 2 @Extended 3 Return 0 = Failed to open a new Writer Document instance.
;                  @Error 2 @Extended 4 Return 0 = Failed to create com.sun.star.sdb.DocumentDefinition Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to insert new Form into Base Document.
;                  @Error 3 @Extended 5 Return 0 = Failed to open new Form Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. New Form was successfully inserted.
;                  @Error 0 @Extended 1 Return Object = Success. Returning opened Form Document's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To create a form inside a folder, the form name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to create FormXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sForm with the following path: Folder1/Folder2/Folder3/FormXYZ.
; Related .......: _LOBase_FormDelete, _LOBase_FormCopy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormCreate(ByRef $oDoc, ByRef $oConnection, $sForm, $bOpen = False, $bDesign = True, $bVisible = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDesktop, $oSource, $oFormDoc, $oDocDef
	Local $asSplit[0]
	Local Const $iURLFrameCreate = 8 ; frame will be created if not found
	Local $aArgs[1]
	Local $iError = 0, $iCount = 0
	Local $sPath = @TempDir & "AutoIt_Form_Temp_Doc_"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOpen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bDesign) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSource = $oDoc.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$asSplit = StringSplit($sForm, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	Next

	$sForm = $asSplit[$asSplit[0]] ; Last element of Array will be the Form name to Create

	If $oSource.hasByName($sForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	$aArgs[0] = __LOBase_SetPropertyValue("Hidden", True)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not IsObj($aArgs[0]) Then $iError = BitOR($iError, 1)
	$oFormDoc = $oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", $iURLFrameCreate, $aArgs)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	$oFormDoc.ApplyFormDesignMode = False

	$oFormDoc.CurrentController.ViewSettings.ShowTableBoundaries = False
	$oFormDoc.CurrentController.ViewSettings.ShowOnlineLayout = True

	While FileExists($sPath & $iCount & ".odt")
		$iCount += 1
		Sleep((IsInt($iCount / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
	WEnd

	ReDim $aArgs[0]

	$sPath &= $iCount & ".odt"
	$oFormDoc.StoreAsUrl(_LOBase_PathConvert($sPath, $LOB_PATHCONV_OFFICE_RETURN), $aArgs)
	$oFormDoc.close(True)

	ReDim $aArgs[3]

	$aArgs[0] = __LOBase_SetPropertyValue("Name", $sForm)
	$aArgs[1] = __LOBase_SetPropertyValue("Parent", $oDoc.FormDocuments())
	$aArgs[2] = __LOBase_SetPropertyValue("URL", _LOBase_PathConvert($sPath, $LOB_PATHCONV_OFFICE_RETURN))

	$oDocDef = $oSource.createInstanceWithArguments("com.sun.star.sdb.DocumentDefinition", $aArgs)
	If Not IsObj($oDocDef) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$oSource.insertbyName($sForm, $oDocDef)

	FileDelete($sPath) ; Delete the file, as it is no longer needed.

	If Not $oSource.hasByName($sForm) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	If $bOpen Then
		If Not $oDoc.CurrentController.isConnected() Then $oDoc.CurrentController.connect()

		If $bDesign Then
			$oFormDoc = $oSource.getByName($sForm).openDesign()

		Else
			$oFormDoc = $oSource.getByName($sForm).open()
		EndIf

		If Not IsObj($oFormDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		$oFormDoc.CurrentController.Frame.ContainerWindow.Visible = $bVisible
		If ($oFormDoc.CurrentController.Frame.ContainerWindow.isVisible() <> $bVisible) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, $oFormDoc)

		Return SetError($__LO_STATUS_SUCCESS, 1, $oFormDoc)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FormCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormDelete
; Description ...: Delete a Form from a Document.
; Syntax ........: _LOBase_FormDelete(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sName               - a string value. The Form name to Delete. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Folder or Sub-Folder not found.
;                  @Error 1 @Extended 4 Return 0 = Name called in $sName not found in Folder.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sName not a Form.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete form.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Form was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To delete a form contained in a folder, you MUST prefix the Form name called in $sName by the folder path it is located in, separated by forward slashes (/). e.g. to delete FormXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/FormXYZ
; Related .......: _LOBase_FormCreate, _LOBase_FormCopy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormDelete(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSource = $oDoc.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Form name to delete

	If Not $oSource.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not $oSource.getByName($sName).supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSource.removeByName($sName)

	If $oSource.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FormDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormDocVisible
; Description ...: Set or retrieve the current visibility of a document.
; Syntax ........: _LOBase_FormDocVisible(ByRef $oFormDoc[, $bVisible = Null])
; Parameters ....: $oFormDoc            - [in/out] an object. A Form Document object returned by a previous _LOBase_FormOpen, _LOBase_FormConnect, or _LOBase_FormCreate function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the document is visible.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormDoc not an Object.
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
Func _LOBase_FormDocVisible(ByRef $oFormDoc, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oFormDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($bVisible = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oFormDoc.CurrentController.Frame.ContainerWindow.isVisible())

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oFormDoc.CurrentController.Frame.ContainerWindow.Visible = $bVisible
	$iError = ($oFormDoc.CurrentController.Frame.ContainerWindow.isVisible() = $bVisible) ? (0) : (1)

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOBase_FormDocVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormExists
; Description ...: Check whether a Document contains a Form by name.
; Syntax ........: _LOBase_FormExists(ByRef $oDoc, $sName[, $bExhaustive = True])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sName               - a string value. The name of the Form to look for. See remarks.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, the search looks inside sub-folders.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Form and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Form or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Form and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended ? Return Boolean = Success. Returning a Boolean value indicating if the Document contains a Form by the called name (True) or not. If True, and $bExhaustive is True, @Extended is set to the number of times a Form with the same name is found in the Document (In sub-folders).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To narrow the search for a form down to a specific folder, you MUST prefix the Form name called in $sName by the folder path to look in, separated by forward slashes (/). e.g. to search for FormXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/FormXYZ
; Related .......: _LOBase_FormDelete, _LOBase_FormOpen, _LOBase_FormsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormExists(ByRef $oDoc, $sName, $bExhaustive = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local $bReturn = False
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iForms = 0
	Local $asNames[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Form name to search

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If $oSource.hasByName($sName) And $oSource.getByName($sName).supportsService("com.sun.star.ucb.Content") Then
		$iForms += 1
		$bReturn = True
	EndIf

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
			ReDim $avFolders[1][2]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				If $avFolders[$iCount][$iObj].hasByName($sName) And $avFolders[$iCount][$iObj].getByName($sName).supportsService("com.sun.star.ucb.Content") Then
					$iForms += 1
					$bReturn = True
				EndIf

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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

	Return SetError($__LO_STATUS_SUCCESS, $iForms, $bReturn)
EndFunc   ;==>_LOBase_FormExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormFolderCopy
; Description ...: Create a copy of an existing Folder.
; Syntax ........: _LOBase_FormFolderCopy(ByRef $oDoc, $sInputFolder, $sOutputFolder)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sInputFolder        - a string value. The Name of the Folder to Copy. Also the Sub-directory the Folder is in. See Remarks.
;                  $sOutputFolder       - a string value. The Name of the Folder to Create. Also the Sub-directory to place the Folder in. See Remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sInputFolder not a String.
;                  @Error 1 @Extended 3 Return 0 = $sOutputFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder name called in $sInputFolder not found.
;                  @Error 1 @Extended 5 Return 0 = Requested Folder not found.
;                  @Error 1 @Extended 6 Return 0 = Folder name called in $sInputFolder not a Folder.
;                  @Error 1 @Extended 7 Return 0 = Folder name called in $sOutputFolder not found.
;                  @Error 1 @Extended 8 Return 0 = Folder already exists with called name in Destination.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.sdb.Forms" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Source Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Folder Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to insert copied Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Copied Folder successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To create a Folder contained in a folder, you MUST prefix the Folder name called in $sInputFolder by the folder path it is located in, separated by forward slashes (/). e.g. to copy FolderXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sInputFolder with the following path: Folder1/Folder2/Folder3/FolderXYZ
;                  To copy a Folder contained in a folder, you MUST prefix the Folder name called in $sOutputFolder by the folder path you want it to be located in, separated by forward slashes (/). e.g. to create FolderXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sOutputFolder with the following path: Folder1/Folder2/Folder3/FolderXYZ
;                  Copying a Folder will copy all contents also.
;                  If only a name is called in $sOutputFolder, the Folder will be created in the main directory, i.e. not inside of any folders.
; Related .......: _LOBase_FormFolderCreate, _LOBase_FormFolderDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormFolderCopy(ByRef $oDoc, $sInputFolder, $sOutputFolder)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oDestination, $oSourceFormFolder, $oFolder
	Local $asSplit[0]
	Local $aArgs[2]
	Local $sSourceFolder, $sDestFolder

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sInputFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sOutputFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sInputFolder, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sSourceFolder = $asSplit[$asSplit[0]] ; Last element of Array will be the Input Form Folder's name to Copy

	If Not $oSource.hasByName($sSourceFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSourceFormFolder = $oSource.getByName($sSourceFolder)
	If Not IsObj($oSourceFormFolder) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If Not $oSourceFormFolder.supportsService("com.sun.star.sdb.Forms") Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oDestination = $oDoc.FormDocuments()
	If Not IsObj($oDestination) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sOutputFolder, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oDestination.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oDestination = $oDestination.getByName($asSplit[$i])
		If Not IsObj($oDestination) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	Next

	$sDestFolder = $asSplit[$asSplit[0]] ; Last element of Array will be the Output Form Folder name to Create

	If $oDestination.hasByName($sDestFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	$aArgs[0] = __LOBase_SetPropertyValue("Name", $sDestFolder)
	$aArgs[1] = __LOBase_SetPropertyValue("EmbeddedObject", $oSourceFormFolder)

	$oFolder = $oDestination.createInstanceWithArguments("com.sun.star.sdb.Forms", $aArgs)
	If Not IsObj($oFolder) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDestination.insertbyName($sDestFolder, $oFolder)
	If Not $oDestination.hasByName($sDestFolder) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FormFolderCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormFolderCreate
; Description ...: Create a Form Folder.
; Syntax ........: _LOBase_FormFolderCreate(ByRef $oDoc, $sFolder)
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
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to insert new Folder into Base Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully created a Folder.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To create a Folder inside a folder, the Folder name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to create FolderXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3/FolderXYZ.
; Related .......: _LOBase_FormFolderCopy, _LOBase_FormFolderDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormFolderCreate(ByRef $oDoc, $sFolder)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oObj
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSource = $oDoc.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sFolder, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sFolder = $asSplit[$asSplit[0]] ; Last element of Array will be the Folder name to Create

	If $oSource.hasByName($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oObj = $oSource.createInstance("com.sun.star.sdb.Forms")

	$oSource.insertbyName($sFolder, $oObj)

	If Not $oSource.hasByName($sFolder) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FormFolderCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormFolderDelete
; Description ...: Delete a Form Folder from a Document.
; Syntax ........: _LOBase_FormFolderDelete(ByRef $oDoc, $sName)
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
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Folder was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To delete a Folder contained in a folder, you MUST prefix the Folder name called in $sName by the folder path it is located in, separated by forward slashes (/). e.g. to delete FolderXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/FolderXYZ
;                  Deleting a Folder will delete all contents also.
; Related .......: _LOBase_FormFolderCopy, _LOBase_FormFolderCreate, _LOBase_FormFoldersGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormFolderDelete(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSource = $oDoc.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Folder name to Delete

	If Not $oSource.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not $oSource.getByName($sName).supportsService("com.sun.star.sdb.Forms") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSource.removeByName($sName)

	If $oSource.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FormFolderDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormFolderExists
; Description ...: Check whether a Document contains a Form Folder by name.
; Syntax ........: _LOBase_FormFolderExists(ByRef $oDoc, $sName[, $bExhaustive = True])
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
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Form and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Form or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Form and Folder names for Sub-Folder.
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
Func _LOBase_FormFolderExists(ByRef $oDoc, $sName, $bExhaustive = True)
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

	$oSource = $oDoc.FormDocuments()
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

	If $oSource.hasByName($sName) And $oSource.getByName($sName).supportsService("com.sun.star.sdb.Forms") Then
		$iResults += 1
		$bReturn = True
	EndIf

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
			ReDim $avFolders[1][2]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				If $avFolders[$iCount][$iObj].hasByName($sName) And $avFolders[$iCount][$iObj].getByName($sName).supportsService("com.sun.star.sdb.Forms") Then
					$iResults += 1
					$bReturn = True
				EndIf

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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
EndFunc   ;==>_LOBase_FormFolderExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormFolderRename
; Description ...: Rename a Form Folder.
; Syntax ........: _LOBase_FormFolderRename(ByRef $oDoc, $sFolder, $sNewName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sFolder             - a string value. The Folder to rename, including the Sub-Folder path, if applicable. See Remarks.
;                  $sNewName            - a string value. The New name to rename the form Folder to.
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
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
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
Func _LOBase_FormFolderRename(ByRef $oDoc, $sFolder, $sNewName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sFolder, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sFolder = $asSplit[$asSplit[0]] ; Last element of Array will be the Folder name to Rename

	If $oSource.hasByName($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If Not $oSource.hasByName($sFolder) Or Not $oSource.getByName($sFolder).supportsService("com.sun.star.sdb.Forms") Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oSource.getByName($sFolder).rename($sNewName)

	If Not $oSource.hasByName($sNewName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FormFolderRename

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormFoldersGetCount
; Description ...: Retrieve a count of Form Folders contained in the Document.
; Syntax ........: _LOBase_FormFoldersGetCount(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
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
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Form and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Form or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Form and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Form Folders contained in the Document as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only the count of main level Folders (not located in folders), or if $bExhaustive is set to True, it will return a count of all Folders contained in the document.
;                  You can narrow the Folder count down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get a count of Folders contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
; Related .......: _LOBase_FormFoldersGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormFoldersGetCount(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
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

	$oSource = $oDoc.FormDocuments()
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

		If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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

					If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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
EndFunc   ;==>_LOBase_FormFoldersGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormFoldersGetNames
; Description ...: Retrieve an array of Folder Names contained in a Document.
; Syntax ........: _LOBase_FormFoldersGetNames(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
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
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Form and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Form or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Form and Folder names for Sub-Folder.
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
; Related .......: _LOBase_FormFolderDelete, _LOBase_FormFolderExists, _LOBase_FormFoldersGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormFoldersGetNames(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
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

	$oSource = $oDoc.FormDocuments()
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

		If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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

					If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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
EndFunc   ;==>_LOBase_FormFoldersGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormIsModified
; Description ...: Test whether the form has been modified since being created or since the last save.
; Syntax ........: _LOBase_FormIsModified(ByRef $oFormDoc)
; Parameters ....: $oFormDoc            - [in/out] an object. A Form Document object returned by a previous _LOBase_FormOpen, _LOBase_FormConnect, or _LOBase_FormCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the Form has been modified since last being saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_FormSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormIsModified(ByRef $oFormDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oFormDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFormDoc.isModified())
EndFunc   ;==>_LOBase_FormIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormOpen
; Description ...: Open a Form Document
; Syntax ........: _LOBase_FormOpen(ByRef $oDoc, ByRef $oConnection, $sName[, $bDesign = True[, $bVisible = True]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sName               - a string value. The Form name to Open. See remarks.
;                  $bDesign             - [optional] a boolean value. Default is True. If True, the form is opened in Design mode.
;                  $bVisible            - [optional] a boolean value. Default is True. If True, the form document will be visible when opened.
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
;                  @Error 1 @Extended 8 Return 0 = Name called in $sName not a Form.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Form Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to open Form Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning opened Form Document's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To open a form located inside a folder, the form name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to open FormXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/FormXYZ.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormOpen(ByRef $oDoc, ByRef $oConnection, $sName, $bDesign = True, $bVisible = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource, $oFormDoc
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDesign) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If Not $oDoc.CurrentController.isConnected() Then $oDoc.CurrentController.connect()

	$oSource = $oDoc.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Form name to Open

	If Not $oSource.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$oObj = $oSource.getByName($sName)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	If Not $oObj.supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	If $bDesign Then
		$oFormDoc = $oObj.openDesign()

	Else
		$oFormDoc = $oObj.open()
	EndIf

	If Not IsObj($oFormDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$oFormDoc.CurrentController.Frame.ContainerWindow.Visible = $bVisible
	If ($oFormDoc.CurrentController.Frame.ContainerWindow.isVisible() <> $bVisible) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, $oFormDoc)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFormDoc)
EndFunc   ;==>_LOBase_FormOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormRename
; Description ...: Rename a Form.
; Syntax ........: _LOBase_FormRename(ByRef $oDoc, $sForm, $sNewName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sForm               - a string value. The Form to rename, including the Sub-Folder path, if applicable. See Remarks.
;                  $sNewName            - a string value. The New name to rename the form to.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sForm not a String.
;                  @Error 1 @Extended 3 Return 0 = $sNewName not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder not found.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sNewName already exists in Folder.
;                  @Error 1 @Extended 6 Return 0 = Form name called in $sForm not found in Folder or is not a Form.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to rename form.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully renamed the Form.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To rename a Form inside a folder, the original Form name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to rename FormXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sForm with the following path: Folder1/Folder2/Folder3/FormXYZ.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormRename(ByRef $oDoc, $sForm, $sNewName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sForm, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sForm = $asSplit[$asSplit[0]] ; Last element of Array will be the Form name to Rename

	If $oSource.hasByName($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If Not $oSource.hasByName($sForm) Or Not $oSource.getByName($sForm).supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oSource.getByName($sForm).rename($sNewName)

	If Not $oSource.hasByName($sNewName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FormRename

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormSave
; Description ...: Save any changes made to a Document.
; Syntax ........: _LOBase_FormSave(ByRef $oFormDoc)
; Parameters ....: $oFormDoc            - [in/out] an object. A Form Document object returned by a previous _LOBase_FormOpen, _LOBase_FormConnect, or _LOBase_FormCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Document called in $oFormDoc has not been saved to a Base Document yet or is read only.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Form Document's Unique Identifier String.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify Form Document's name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Array of Form and Folder names.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Form or Folder Object.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Array of Form and Folder names for Sub-Folder.
;                  @Error 3 @Extended 7 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  @Error 3 @Extended 8 Return 0 = Failed to identify Form in Parent Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Form was successfully saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_FormIsModified
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormSave(ByRef $oFormDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oForm, $oObj
	Local $asResult[0], $asNames[0], $asFolderList[0]
	Local $iCount = 0, $iFolders = 1
	Local $avFolders[0][2]
	Local $sUniqueString, $sTitle
	Local Enum $iName, $iObj, $iPrefix
	Local Const $__STR_REGEXPARRAYMATCH = 1

	If Not IsObj($oFormDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oFormDoc.hasLocation Or $oFormDoc.isReadOnly Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	; Retrieve the unique name of the form.
	; file:///C:/Folder1/Autoit/Libre%20Office%20Base%20Info/Testing.odb/Obj61/
	$asResult = StringRegExp($oFormDoc.NameSpace(), "/([a-zA-Z0-9]+)/$", $__STR_REGEXPARRAYMATCH)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sUniqueString = $asResult[0]

	; "Testing.odb : abc1Form2"
	$asResult = StringRegExp($oFormDoc.Title(), "\: (.+)$", $__STR_REGEXPARRAYMATCH)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$sTitle = $asResult[0]

	$oSource = $oFormDoc.Parent.FormDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	If $oSource.hasByName($sTitle) And $oSource.getByName($sTitle).supportsService("com.sun.star.ucb.Content") And ($oSource.getByName($sTitle).PersistentName() = $sUniqueString) Then
		$oForm = $oSource.getByName($sTitle)

	Else

		For $i = 0 To UBound($asNames) - 1
			$oObj = $oSource.getByName($asNames[$i])
			If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

			If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
				ReDim $avFolders[1][2]
				$avFolders[0][$iName] = $asNames[$i]
				$avFolders[0][$iObj] = $oObj
			EndIf

			While ($iCount < UBound($avFolders))
				If $avFolders[$iCount][$iObj].hasByName($sTitle) And $avFolders[$iCount][$iObj].getByName($sTitle).supportsService("com.sun.star.ucb.Content") And _
						($avFolders[$iCount][$iObj].getByName($sTitle).PersistentName() = $sUniqueString) Then
					$oForm = $avFolders[$iCount][$iObj].getByName($sTitle)
					ExitLoop
				EndIf

				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

					If $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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

	If Not IsObj($oForm) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

	$oForm.Store()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FormSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormsGetCount
; Description ...: Retrieve a count of Forms contained in the Document.
; Syntax ........: _LOBase_FormsGetCount(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, retrieves a count of all forms, including those in sub-folders.
;                  $sFolder             - [optional] a string value. Default is "". The Folder to return the count of forms for. See remarks.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder called in $sFolder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Form and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Form or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Form and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Forms contained in the Document as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only the count of main level Forms (not located in folders), or if $bExhaustive is set to True, the return will be a count of all forms contained in the document.
;                  You can narrow the Form count down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get a count of forms contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
; Related .......: _LOBase_FormsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormsGetCount(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iForms = 0
	Local $asNames[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.FormDocuments()
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

		If $oObj.supportsService("com.sun.star.ucb.Content") Then ; Form Doc.
			$iForms += 1

		ElseIf $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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

					If $oObj.supportsService("com.sun.star.ucb.Content") Then ; Form Doc.
						$iForms += 1

					ElseIf $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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

	Return SetError($__LO_STATUS_SUCCESS, 0, $iForms)
EndFunc   ;==>_LOBase_FormsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormsGetNames
; Description ...: Retrieve an Array of Form Names contained in a Document.
; Syntax ........: _LOBase_FormsGetNames(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, retrieves all form names, including those in sub-folders.
;                  $sFolder             - [optional] a string value. Default is "". The Sub-Folder to return the array of form names from. See remarks.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder called in $sFolder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Form and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Form or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Form and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Form names contained in this Document. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only an array of main level Form names (not located in folders), or if $bExhaustive is set to True, it will return an array of all forms contained in the document.
;                  You can narrow the Form name return down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get an array of forms contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
;                  All forms located in folders will have the folder path prefixed to the Form name, separated by forward slashes (/). e.g. Folder1/Folder2/Folder3/FormXYZ.
;                  Calling $bExhaustive with True when searching inside a Folder, will get all Form names from inside that folder, and all sub-folders.
;                  The order of the form names inside the folders may not necessarily be in proper order, i.e. if there are two sub folders, and folders inside the first sub-folder, the Forms inside of the two folders will be listed first, then the forms inside the folders inside the first sub-folder.
; Related .......: _LOBase_FormsGetCount, _LOBase_FormDelete, _LOBase_FormOpen
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormsGetNames(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iForms = 0
	Local $asNames[0], $asForms[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.FormDocuments()
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
			If (UBound($asForms) >= $iForms) Then ReDim $asForms[$iForms + 1]
			$asForms[$iForms] = $sFolder & $asNames[$i]
			$iForms += 1

		ElseIf $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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
						If (UBound($asForms) >= $iForms) Then ReDim $asForms[$iForms + 1]
						$asForms[$iForms] = $sFolder & $avFolders[$iCount][$iPrefix] & $asFolderList[$k]
						$iForms += 1

					ElseIf $oObj.supportsService("com.sun.star.sdb.Forms") Then ; Folder.
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

	Return SetError($__LO_STATUS_SUCCESS, UBound($asForms), $asForms)
EndFunc   ;==>_LOBase_FormsGetNames
