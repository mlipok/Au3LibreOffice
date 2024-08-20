#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;~ #Tidy_Parameters=/sf
#include-once
#include "LibreOffice_Constants.au3"

; Common includes for Base
#include "LibreOfficeBase_Internal.au3"

; Other includes for Base

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Registering, Unregistering, and connecting to, etc. L.O. Base Databases.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_DatabaseConnectionClose
; _LOBase_DatabaseConnectionGet
; _LOBase_DatabaseGetObjByDoc
; _LOBase_DatabaseGetObjByURL
; _LOBase_DatabaseIsReadOnly
; _LOBase_DatabaseName
; _LOBase_DatabaseRegisteredAdd
; _LOBase_DatabaseRegisteredExists
; _LOBase_DatabaseRegisteredGetNames
; _LOBase_DatabaseRegisteredRemoveByName
; _LOBase_DatabaseRequiresPassword
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseConnectionClose
; Description ...: Close an opened Database connection.
; Syntax ........: _LOBase_DatabaseConnectionClose(ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a connection Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection is already closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to close connection.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully closed the connection.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseConnectionGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseConnectionClose(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oConnection.ImplementationName() <> "com.sun.star.sdbc.drivers.OConnectionWrapper" Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oConnection.Close()

	If Not $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DatabaseConnectionClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseConnectionGet
; Description ...: Create a connection to a Database.
; Syntax ........: _LOBase_DatabaseConnectionGet(ByRef $oDBase[, $sUser = ""[, $sPass = ""[, $bPromptUser = False]]])
; Parameters ....: $oDBase              - [in/out] an object. A Database object returned by a previous _LOBase_DatabaseGetObjByDoc or _LOBase_DatabaseGetObjByURL function.
;                  $sUser               - [optional] a string value. Default is "". The Username for connecting to the Database. If none, leave as a blank string.
;                  $sPass               - [optional] a string value. Default is "". The Password for connecting to the Database. If none, leave as a blank string.
;                  $bPromptUser         - [optional] a boolean value. Default is False. If True, $sUser and $sPass are ignored, and the user is prompted for the required information.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDBase not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sUser not a String.
;                  @Error 1 @Extended 3 Return 0 = $sPass not a String.
;                  @Error 1 @Extended 4 Return 0 = $bPromptUser not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.task.InteractionHandler" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Connect to the Database.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully connected to the Database, returning a Connection Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseConnectionClose
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseConnectionGet(ByRef $oDBase, $sUser = "", $sPass = "", $bPromptUser = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBConnection, $oHandler

	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sUser) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sPass) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bPromptUser) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If $bPromptUser Then
		$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oHandler = $oServiceManager.createInstance("com.sun.star.task.InteractionHandler")
		If Not IsObj($oHandler) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oDBConnection = $oDBase.ConnectWithCompletion($oHandler)

	Else
		$oDBConnection = $oDBase.getConnection($sUser, $sPass)

	EndIf

	If Not IsObj($oDBConnection) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDBConnection)
EndFunc   ;==>_LOBase_DatabaseConnectionGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseGetObjByDoc
; Description ...: Retrieve a Database Object from a Document Object.
; Syntax ........: _LOBase_DatabaseGetObjByDoc(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Document Save path.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Database Object.
;                  --Document Errors--
;                  @Error 5 @Extended 1 Return 0 = Document hasn't been saved yet.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Database Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Document must be saved first, before you can obtain a Database Object for it.
;                  If you intend to move or delete the Database Document file after retrieving a Database Object, you must allow the function that created the Database Object to return, and the Database Object to be released by AutoIt (only works for Local variable, by the way). Otherwise the file will be shown as "In Use" and cannot be moved or deleted until AutoIt releases the Object. See AutoIt Help file "Obj/COM Reference", Last paragraph in section "An example usage of COM in AutoIt"
; Related .......: _LOBase_DatabaseGetObjByURL
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseGetObjByDoc(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext, $oDBase
	Local $sURL

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sURL = $oDoc.URL()
	If Not IsString($sURL) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sURL = "") Then Return SetError($__LO_STATUS_DOC_ERROR, 1, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oDBase = $oDBaseContext.getByName($sURL)
	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDBase)
EndFunc   ;==>_LOBase_DatabaseGetObjByDoc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseGetObjByURL
; Description ...: Retrieve a Database Object from a URL or registered Database name.
; Syntax ........: _LOBase_DatabaseGetObjByURL($sURL)
; Parameters ....: $sURL                - a string value. The File path of the Database file or a Database name that is registered.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sURL not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to convert called path to URL format.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Database Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Database Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving a Database Object this way allows you to edit a Database without using a GUI.
;                  If you intend to move or delete the Database Document file after retrieving a Database Object, you must allow the function that created the Database Object to return, and the Database Object to be released by AutoIt (only works for Local variable, by the way). Otherwise the file will be shown as "In Use" and cannot be moved or deleted until AutoIt releases the Object. See AutoIt Help file "Obj/COM Reference", Last paragraph in section "An example usage of COM in AutoIt"
; Related .......: _LOBase_DatabaseGetObjByDoc
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseGetObjByURL($sURL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext, $oDBase
	Local $sFileURL = $sURL

	If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If StringInStr($sFileURL, "\") Then $sFileURL = _LOBase_PathConvert($sFileURL, $LOB_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oDBase = $oDBaseContext.getByName($sFileURL)
	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDBase)
EndFunc   ;==>_LOBase_DatabaseGetObjByURL

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseIsReadOnly
; Description ...: Check if a Database is currently in a Read-Only state.
; Syntax ........: _LOBase_DatabaseIsReadOnly(ByRef $oDBase)
; Parameters ....: $oDBase              - [in/out] an object. A Database object returned by a previous _LOBase_DatabaseGetObjByDoc or _LOBase_DatabaseGetObjByURL function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDBase not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query Database for Read-Only Status.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If Database is currently Read-Only, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseRequiresPassword
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseIsReadOnly(ByRef $oDBase)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn

	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bReturn = $oDBase.DatabaseDocument.DataSource.IsReadOnly()
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_DatabaseIsReadOnly

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseName
; Description ...: Retrieve the Database Name property. See remarks.
; Syntax ........: _LOBase_DatabaseName(ByRef $oDBase)
; Parameters ....: $oDBase              - [in/out] an object.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDBase not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Database Source name.
;                  @Error 3 @Extended 2 Return 0 = Failed to convert URL to Computer Path.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the Name value as a String. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the database data source is registered, then the Name property denotes the registration name. Otherwise, the name property contains the URL of the file.
;                  If the same database data source is registered under different names, the value of the Name property is not defined.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseName(ByRef $oDBase)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sName = $oDBase.DatabaseDocument.DataSource.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If StringInStr($sName, "/") Then $sName = _LOBase_PathConvert($sName, $LOB_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sName)
EndFunc   ;==>_LOBase_DatabaseName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRegisteredAdd
; Description ...: Add a Database to the Registered Databases list.
; Syntax ........: _LOBase_DatabaseRegisteredAdd(ByRef $oDBase, $sName)
; Parameters ....: $oDBase              - [in/out] an object. A Database object returned by a previous _LOBase_DatabaseGetObjByDoc or _LOBase_DatabaseGetObjByURL function.
;                  $sName               - a string value. The name to register the Database under.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDBase not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Registered Database already exists with name as called in $sName.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to register Database.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Database successfully registered.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseRegisteredExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRegisteredAdd(ByRef $oDBase, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext

	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If $oDBaseContext.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oDBaseContext.registerObject($sName, $oDBase)
	If Not $oDBaseContext.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DatabaseRegisteredAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRegisteredExists
; Description ...: Query whether a Registered Database exists by name.
; Syntax ........: _LOBase_DatabaseRegisteredExists($sName)
; Parameters ....: $sName               - a string value. The Database name to look for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sName not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query for Registered Database.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If Registered Database with called name exists, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRegisteredExists($sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext
	Local $bReturn

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$bReturn = $oDBaseContext.hasByName($sName)
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_DatabaseRegisteredExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRegisteredGetNames
; Description ...: Retrieve an array of Registered Database names.
; Syntax ........: _LOBase_DatabaseRegisteredGetNames()
; Parameters ....: None
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Registered Database Names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning array of Registered Database names. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseGetObjByURL
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRegisteredGetNames()
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext
	Local $asNames[0]

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$asNames = $oDBaseContext.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOBase_DatabaseRegisteredGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRegisteredRemoveByName
; Description ...: Remove a registered Database by name.
; Syntax ........: _LOBase_DatabaseRegisteredRemoveByName($sName)
; Parameters ....: $sName               - a string value. The Registered Database name to remove.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sName not a String.
;                  @Error 1 @Extended 2 Return 0 = No Registered Database found with name as called in $sName.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to unregister Database.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Database successfully unregistered.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseRegisteredAdd, _LOBase_DatabaseRegisteredExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRegisteredRemoveByName($sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not $oDBaseContext.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDBaseContext.revokeObject($sName)
	If $oDBaseContext.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DatabaseRegisteredRemoveByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRequiresPassword
; Description ...: Check if a Database requires a password in order to connect to it.
; Syntax ........: _LOBase_DatabaseRequiresPassword(ByRef $oDBase)
; Parameters ....: $oDBase              - [in/out] an object. A Database object returned by a previous _LOBase_DatabaseGetObjByDoc or _LOBase_DatabaseGetObjByURL function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDBase not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query Database for password requirement.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If Database requires a password to connect to it, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRequiresPassword(ByRef $oDBase)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn

	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bReturn = $oDBase.DatabaseDocument.DataSource.IsPasswordRequired()
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_DatabaseRequiresPassword
