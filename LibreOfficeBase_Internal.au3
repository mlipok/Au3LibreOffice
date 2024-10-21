#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;~ #Tidy_Parameters=/sf
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Base
#include "LibreOfficeBase_Constants.au3"
#include "LibreOfficeBase_Helper.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Various functions for internal data processing, data retrieval, retrieving and applying settings for LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __LOBase_AddTo1DArray
; __LOBase_ArrayFill
; __LOBase_ColTransferProps
; __LOBase_ColTypeName
; __LOBase_CreateStruct
; __LOBase_DatabaseMetaGetQuery
; __LOBase_InternalComErrorHandler
; __LOBase_IntIsBetween
; __LOBase_SetPropertyValue
; __LOBase_VarsAreNull
; __LOBase_VersionCheck
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_AddTo1DArray
; Description ...: Add data to a 1 Dimensional array.
; Syntax ........: __LOBase_AddTo1DArray(ByRef $aArray, $vData[, $bCountInFirst = False])
; Parameters ....: $aArray              - [in/out] an array of unknowns. The Array to directly add data to. Array will be directly modified.
;                  $vData               - a variant value. The Data to add to the Array.
;                  $bCountInFirst       - [optional] a boolean value. Default is False. If True the first element of the array is a count of contained elements.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $aArray not an Array
;                  @Error 1 @Extended 2 Return 0 = $bCountinFirst not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $aArray contains too many columns.
;                  @Error 1 @Extended 4 Return 0 = $aArray[0] contains non integer data or is not empty, and $bCountInFirst is set to True.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Array item was successfully added.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_AddTo1DArray(ByRef $aArray, $vData, $bCountInFirst = False)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($aArray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bCountInFirst) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If UBound($aArray, $UBOUND_COLUMNS) > 1 Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Too many columns

	If $bCountInFirst And (UBound($aArray) = 0) Then
		ReDim $aArray[1]
		$aArray[0] = 0
	EndIf

	If $bCountInFirst And (($aArray[0] <> "") And Not IsInt($aArray[0])) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	ReDim $aArray[UBound($aArray) + 1]
	$aArray[UBound($aArray) - 1] = $vData
	If $bCountInFirst Then $aArray[0] += 1
	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOBase_AddTo1DArray

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_ArrayFill
; Description ...: Fill an Array with data.
; Syntax ........: __LOBase_ArrayFill(ByRef $aArrayToFill[, $vVar1 = Null[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null[, $vVar9 = Null[, $vVar10 = Null[, $vVar11 = Null[, $vVar12 = Null[, $vVar13 = Null[, $vVar14 = Null[, $vVar15 = Null[, $vVar16 = Null[, $vVar17 = Null[, $vVar18 = Null]]]]]]]]]]]]]]]]]])
; Parameters ....: $aArrayToFill        - [in/out] an array of unknowns. The Array to Fill. Array will be directly modified.
;                  $vVar1               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar2               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar3               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar4               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar5               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar6               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar7               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar8               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar9               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar10              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar11              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar12              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar13              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar14              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar15              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar16              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar17              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar18              - [optional] a variant value. Default is Null. The Data to add to the Array.
; Return values .: None
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call only how many you parameters you need to add to the Array. Automatically resizes the Array if it is the incorrect size.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_ArrayFill(ByRef $aArrayToFill, $vVar1 = Null, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null, $vVar9 = Null, $vVar10 = Null, $vVar11 = Null, $vVar12 = Null, $vVar13 = Null, $vVar14 = Null, $vVar15 = Null, $vVar16 = Null, $vVar17 = Null, $vVar18 = Null)
	#forceref $vVar1, $vVar2, $vVar3, $vVar4, $vVar5, $vVar6, $vVar7, $vVar8, $vVar9, $vVar10, $vVar11, $vVar12, $vVar13, $vVar14, $vVar15, $vVar16, $vVar17, $vVar18

	If UBound($aArrayToFill) < (@NumParams - 1) Then ReDim $aArrayToFill[@NumParams - 1]
	For $i = 0 To @NumParams - 2
		$aArrayToFill[$i] = Eval("vVar" & $i + 1)
	Next
EndFunc   ;==>__LOBase_ArrayFill

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_ColTransferProps
; Description ...: Transfer column properties from one to another.
; Syntax ........: __LOBase_ColTransferProps(ByRef $oNewCol, ByRef $oOldCol)
; Parameters ....: $oNewCol             - [in/out] an object. A new column Object.
;                  $oOldCol             - [in/out] an object. A Column object returned by a previous _LOBase_TableColGetObjByIndex or _LOBase_TableColGetObjByName function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oNewCol not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oOldCol not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Old Column's Properties.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully transferred Column properties.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_ColTransferProps(ByRef $oNewCol, ByRef $oOldCol)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atProperties[0]

	If Not IsObj($oNewCol) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oOldCol) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$atProperties = $oOldCol.getPropertySetInfo.Properties()
	If Not IsArray($atProperties) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($atProperties) - 1
		If ($oOldCol.getPropertyValue($atProperties[$i].Name()) <> "") Then $oNewCol.setPropertyValue($atProperties[$i].Name(), $oOldCol.getPropertyValue($atProperties[$i].Name()))
		Sleep(($i = $__LOBCONST_SLEEP_DIV) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOBase_ColTransferProps

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_ColTypeName
; Description ...: Obtain an appropriate Type Name for a Column Type.
; Syntax ........: __LOBase_ColTypeName($iType)
; Parameters ....: $iType               - an integer value. The Column Type. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iType not an integer, less than -16 or greater than 2014. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 2 Return 0 = $iType not one of the pre-defined constants.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the Type name corresponding to the Type Constant.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_ColTypeName($iType)
	Local $sType

	If Not __LOBase_IntIsBetween($iType, $LOB_DATA_TYPE_LONGNVARCHAR, $LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $iType

		Case $LOB_DATA_TYPE_LONGNVARCHAR
			$sType = "LONGNVARCHAR"

		Case $LOB_DATA_TYPE_NCHAR
			$sType = "NCHAR"

		Case $LOB_DATA_TYPE_NVARCHAR
			$sType = "NVARCHAR"

		Case $LOB_DATA_TYPE_ROWID
			$sType = "ROWID"

		Case $LOB_DATA_TYPE_BIT
			$sType = "BIT"

		Case $LOB_DATA_TYPE_TINYINT
			$sType = "TINYINT"

		Case $LOB_DATA_TYPE_BIGINT
			$sType = "BIGINT"

		Case $LOB_DATA_TYPE_LONGVARBINARY
			$sType = "LONGVARBINARY"

		Case $LOB_DATA_TYPE_VARBINARY
			$sType = "VARBINARY"

		Case $LOB_DATA_TYPE_BINARY
			$sType = "BINARY"

		Case $LOB_DATA_TYPE_LONGVARCHAR
			$sType = "LONGVARCHAR"

		Case $LOB_DATA_TYPE_SQLNULL
			$sType = "SQLNULL"

		Case $LOB_DATA_TYPE_CHAR
			$sType = "CHAR"

		Case $LOB_DATA_TYPE_NUMERIC
			$sType = "NUMERIC"

		Case $LOB_DATA_TYPE_DECIMAL
			$sType = "DECIMAL"

		Case $LOB_DATA_TYPE_INTEGER
			$sType = "INTEGER"

		Case $LOB_DATA_TYPE_SMALLINT
			$sType = "SMALLINT"

		Case $LOB_DATA_TYPE_FLOAT
			$sType = "FLOAT"

		Case $LOB_DATA_TYPE_REAL
			$sType = "REAL"

		Case $LOB_DATA_TYPE_DOUBLE
			$sType = "DOUBLE"

		Case $LOB_DATA_TYPE_VARCHAR
			$sType = "VARCHAR"

		Case $LOB_DATA_TYPE_BOOLEAN
			$sType = "BOOLEAN"

		Case $LOB_DATA_TYPE_DATALINK
			$sType = "DATALINK"

		Case $LOB_DATA_TYPE_DATE
			$sType = "DATE"

		Case $LOB_DATA_TYPE_TIME
			$sType = "TIME"

		Case $LOB_DATA_TYPE_TIMESTAMP
			$sType = "TIMESTAMP"

		Case $LOB_DATA_TYPE_OTHER
			$sType = "OTHER"

		Case $LOB_DATA_TYPE_OBJECT
			$sType = "OBJECT"

		Case $LOB_DATA_TYPE_DISTINCT
			$sType = "DISTINCT"

		Case $LOB_DATA_TYPE_STRUCT
			$sType = "STRUCT"

		Case $LOB_DATA_TYPE_ARRAY
			$sType = "ARRAY"

		Case $LOB_DATA_TYPE_BLOB
			$sType = "BLOB"

		Case $LOB_DATA_TYPE_CLOB
			$sType = "CLOB"

		Case $LOB_DATA_TYPE_REF
			$sType = "REF"

		Case $LOB_DATA_TYPE_SQLXML
			$sType = "SQLXML"

		Case $LOB_DATA_TYPE_NCLOB
			$sType = "NCLOB"

		Case $LOB_DATA_TYPE_REF_CURSOR
			$sType = "REF_CURSOR"

		Case $LOB_DATA_TYPE_TIME_WITH_TIMEZONE
			$sType = "TIME_WITH_TIMEZONE"

		Case $LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE
			$sType = "TIMESTAMP_WITH_TIMEZONE"

		Case Else
			Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, $sType)
EndFunc   ;==>__LOBase_ColTypeName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_CreateStruct
; Description ...: Creates a Struct.
; Syntax ........: __LOBase_CreateStruct($sStructName)
; Parameters ....: $sStructName         - a string value. Name of structure to create.
; Return values .: Success: Structure.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sStructName not a string
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object
;                  @Error 2 @Extended 2 Return 0 = Error creating requested structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Structure = Success. Property Structure Returned
; Author ........: mLipok
; Modified ......: donnyh13 - Added error checking.
; Remarks .......: From WriterDemo.au3 as modified by mLipok from WriterDemo.vbs found in the LibreOffice SDK examples.
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/204665-libreopenoffice-writer/?do=findComment&comment=1471711
; Example .......: No
; ===============================================================================================================================
Func __LOBase_CreateStruct($sStructName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $tStruct

	If Not IsString($sStructName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$tStruct = $oServiceManager.Bridge_GetStruct($sStructName)
	If Not IsObj($tStruct) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $tStruct)
EndFunc   ;==>__LOBase_CreateStruct

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_DatabaseMetaGetQuery
; Description ...: Return the Query command from a Constant value.
; Syntax ........: __LOBase_DatabaseMetaGetQuery($iQuery)
; Parameters ....: $iQuery              - an integer value. The Query to retrieve the command for. See Constants, $LOB_DBASE_META_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iQuery not an Integer, less than 0 or greater than number of query commands. See Constants, $LOB_DBASE_META_* as defined in LibreOfficeBase_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the requested Query command.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_DatabaseMetaGetQuery($iQuery)
	Local $asMetaQueries[148]

	If Not __LOBase_IntIsBetween($iQuery, 0, UBound($asMetaQueries)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asMetaQueries[$LOB_DBASE_META_ALL_PROCEDURES_ARE_CALLABLE] = ".allProceduresAreCallable"
	$asMetaQueries[$LOB_DBASE_META_ALL_TABLES_ARE_SELECTABLE] = ".allTablesAreSelectable"
	$asMetaQueries[$LOB_DBASE_META_DATA_DEFINITION_CAUSES_TRANSACTION_COMMIT] = ".dataDefinitionCausesTransactionCommit"
	$asMetaQueries[$LOB_DBASE_META_DATA_DEFINITION_IGNORED_IN_TRANSACTIONS] = ".dataDefinitionIgnoredInTransactions"
	$asMetaQueries[$LOB_DBASE_META_DELETES_ARE_DETECTED] = ".deletesAreDetected"
	$asMetaQueries[$LOB_DBASE_META_DOES_MAX_ROW_SIZE_INCLUDE_BLOBS] = ".doesMaxRowSizeIncludeBlobs"
	$asMetaQueries[$LOB_DBASE_META_GET_BEST_ROW_ID] = ".getBestRowIdentifier"
	$asMetaQueries[$LOB_DBASE_META_GET_CATALOG_SEPARATOR] = ".getCatalogSeparator"
	$asMetaQueries[$LOB_DBASE_META_GET_CATALOG_TERM] = ".getCatalogTerm"
	$asMetaQueries[$LOB_DBASE_META_GET_CATALOGS] = ".getCatalogs"
	$asMetaQueries[$LOB_DBASE_META_GET_COLS] = ".getColumns"
	$asMetaQueries[$LOB_DBASE_META_GET_COL_PRIVILEGES] = ".getColumnPrivileges"
	$asMetaQueries[$LOB_DBASE_META_GET_CROSS_REF] = ".getCrossReference"
	$asMetaQueries[$LOB_DBASE_META_GET_DATABASE_PRODUCT_NAME] = ".getDatabaseProductName"
	$asMetaQueries[$LOB_DBASE_META_GET_DATABASE_PRODUCT_VERSION] = ".getDatabaseProductVersion"
	$asMetaQueries[$LOB_DBASE_META_GET_DEFAULT_TRANSACTION_ISOLATION] = ".getDefaultTransactionIsolation"
	$asMetaQueries[$LOB_DBASE_META_GET_DRIVER_MAJOR_VERSION] = ".getDriverMajorVersion"
	$asMetaQueries[$LOB_DBASE_META_GET_DRIVER_MINOR_VERSION] = ".getDriverMinorVersion"
	$asMetaQueries[$LOB_DBASE_META_GET_DRIVER_NAME] = ".getDriverName"
	$asMetaQueries[$LOB_DBASE_META_GET_DRIVER_VERSION] = ".getDriverVersion"
	$asMetaQueries[$LOB_DBASE_META_GET_EXPORTED_KEYS] = ".getExportedKeys"
	$asMetaQueries[$LOB_DBASE_META_GET_EXTRA_NAME_CHARS] = ".getExtraNameCharacters"
	$asMetaQueries[$LOB_DBASE_META_GET_IDENTIFIER_QUOTE_STRING] = ".getIdentifierQuoteString"
	$asMetaQueries[$LOB_DBASE_META_GET_IMPORTED_KEYS] = ".getImportedKeys"
	$asMetaQueries[$LOB_DBASE_META_GET_INDEX_INFO] = ".getIndexInfo"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_BINARY_LITERAL_LEN] = ".getMaxBinaryLiteralLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_CATALOG_NAME_LEN] = ".getMaxCatalogNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_CHAR_LITERAL_LEN] = ".getMaxCharLiteralLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COL_NAME_LEN] = ".getMaxColumnNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COLS_IN_GROUP_BY] = ".getMaxColumnsInGroupBy"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COLS_IN_INDEX] = ".getMaxColumnsInIndex"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COLS_IN_ORDER_BY] = ".getMaxColumnsInOrderBy"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COLS_IN_SEL] = ".getMaxColumnsInSelect"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COLS_IN_TABLE] = ".getMaxColumnsInTable"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_CONNECTIONS] = ".getMaxConnections"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_CURSOR_NAME_LEN] = ".getMaxCursorNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_INDEX_LEN] = ".getMaxIndexLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_PROCEDURE_NAME_LEN] = ".getMaxProcedureNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_ROW_SIZE] = ".getMaxRowSize"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_SCHEMA_NAME_LEN] = ".getMaxSchemaNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_STATEMENT_LEN] = ".getMaxStatementLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_STATEMENTS] = ".getMaxStatements"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_TABLE_NAME_LEN] = ".getMaxTableNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_TABLES_IN_SEL] = ".getMaxTablesInSelect"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_USER_NAME_LEN] = ".getMaxUserNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_NUMERIC_FUNCS] = ".getNumericFunctions"
	$asMetaQueries[$LOB_DBASE_META_GET_PRIMARY_KEY] = ".getPrimaryKeys"
	$asMetaQueries[$LOB_DBASE_META_GET_PROCEDURE_COLS] = ".getProcedureColumns"
	$asMetaQueries[$LOB_DBASE_META_GET_PROCEDURE_TERM] = ".getProcedureTerm"
	$asMetaQueries[$LOB_DBASE_META_GET_PROCEDURES] = ".getProcedures"
	$asMetaQueries[$LOB_DBASE_META_GET_SCHEMA_TERM] = ".getSchemaTerm"
	$asMetaQueries[$LOB_DBASE_META_GET_SCHEMAS] = ".getSchemas"
	$asMetaQueries[$LOB_DBASE_META_GET_SEARCH_STRING_ESCAPE] = ".getSearchStringEscape"
	$asMetaQueries[$LOB_DBASE_META_GET_SQL_KEYWORDS] = ".getSQLKeywords"
	$asMetaQueries[$LOB_DBASE_META_GET_STRING_FUNCS] = ".getStringFunctions"
	$asMetaQueries[$LOB_DBASE_META_GET_SYSTEM_FUNCS] = ".getSystemFunctions"
	$asMetaQueries[$LOB_DBASE_META_GET_TABLE_PRIVILEGES] = ".getTablePrivileges"
	$asMetaQueries[$LOB_DBASE_META_GET_TABLE_TYPES] = ".getTableTypes"
	$asMetaQueries[$LOB_DBASE_META_GET_TABLES] = ".getTables"
	$asMetaQueries[$LOB_DBASE_META_GET_TIME_DATE_FUNCS] = ".getTimeDateFunctions"
	$asMetaQueries[$LOB_DBASE_META_GET_TYPE_INFO] = ".getTypeInfo"
	$asMetaQueries[$LOB_DBASE_META_GET_UDTS] = ".getUDTs"
	$asMetaQueries[$LOB_DBASE_META_GET_URL] = ".getURL"
	$asMetaQueries[$LOB_DBASE_META_GET_USERNAME] = ".getUserName"
	$asMetaQueries[$LOB_DBASE_META_GET_VERSION_COLS] = ".getVersionColumns"
	$asMetaQueries[$LOB_DBASE_META_INSERTS_ARE_DETECTED] = ".insertsAreDetected"
	$asMetaQueries[$LOB_DBASE_META_IS_CATALOG_AT_START] = ".isCatalogAtStart"
	$asMetaQueries[$LOB_DBASE_META_IS_READ_ONLY] = ".isReadOnly"
	$asMetaQueries[$LOB_DBASE_META_NULL_PLUS_NON_NULL_IS_NULL] = ".nullPlusNonNullIsNull"
	$asMetaQueries[$LOB_DBASE_META_NULLS_ARE_SORTED_AT_END] = ".nullsAreSortedAtEnd"
	$asMetaQueries[$LOB_DBASE_META_NULLS_ARE_SORTED_AT_START] = ".nullsAreSortedAtStart"
	$asMetaQueries[$LOB_DBASE_META_NULLS_ARE_SORTED_HIGH] = ".nullsAreSortedHigh"
	$asMetaQueries[$LOB_DBASE_META_NULLS_ARE_SORTED_LOW] = ".nullsAreSortedLow"
	$asMetaQueries[$LOB_DBASE_META_OTHERS_DELETES_ARE_VISIBLE] = ".othersDeletesAreVisible"
	$asMetaQueries[$LOB_DBASE_META_OTHERS_INSERTS_ARE_VISIBLE] = ".othersInsertsAreVisible"
	$asMetaQueries[$LOB_DBASE_META_OTHERS_UPDATES_ARE_VISIBLE] = ".othersUpdatesAreVisible"
	$asMetaQueries[$LOB_DBASE_META_OWN_DELETES_ARE_VISIBLE] = ".ownDeletesAreVisible"
	$asMetaQueries[$LOB_DBASE_META_OWN_INSERTS_ARE_VISIBLE] = ".ownInsertsAreVisible"
	$asMetaQueries[$LOB_DBASE_META_OWN_UPDATES_ARE_VISIBLE] = ".ownUpdatesAreVisible"
	$asMetaQueries[$LOB_DBASE_META_STORES_LOWER_CASE_IDS] = ".storesLowerCaseIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_STORES_MIXED_CASE_IDS] = ".storesMixedCaseIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_STORES_UPPER_CASE_IDS] = ".storesUpperCaseIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_STORES_LOWER_CASE_QUOTED_IDS] = ".storesLowerCaseQuotedIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_STORES_MIXED_CASE_QUOTED_IDS] = ".storesMixedCaseQuotedIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_STORES_UPPER_CASE_QUOTED_IDS] = ".storesUpperCaseQuotedIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ALTER_TABLE_WITH_ADD_COL] = ".supportsAlterTableWithAddColumn"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ALTER_TABLE_WITH_DROP_COL] = ".supportsAlterTableWithDropColumn"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ANSI92_ENTRY_LEVEL_SQL] = ".supportsANSI92EntryLevelSQL"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ANSI92_FULL_SQL] = ".supportsANSI92FullSQL"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ANSI92_INTERMEDIATE_SQL] = ".supportsANSI92IntermediateSQL"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_BATCH_UPDATES] = ".supportsBatchUpdates"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_DATA_MANIPULATION] = ".supportsCatalogsInDataManipulation"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_INDEX_DEFINITIONS] = ".supportsCatalogsInIndexDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_PRIVILEGE_DEFINITIONS] = ".supportsCatalogsInPrivilegeDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_PROCEDURE_CALLS] = ".supportsCatalogsInProcedureCalls"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_TABLE_DEFINITION] = ".supportsCatalogsInTableDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_COL_ALIASING] = ".supportsColumnAliasing"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CONVERT] = ".supportsConvert"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CORE_SQL_GRAMMAR] = ".supportsCoreSQLGrammar"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CORRELATED_SUBQUERIES] = ".supportsCorrelatedSubqueries"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_DATA_DEFINITION_AND_DATA_MANIPULATION_TRANSACTIONS] = ".supportsDataDefinitionAndDataManipulationTransactions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_DATA_MANIPULATION_TRANSACTIONS_ONLY] = ".supportsDataManipulationTransactionsOnly"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_DIFF_TABLE_CORRELATION_NAMES] = ".supportsDifferentTableCorrelationNames"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_EXPRESSIONS_IN_ORDER_BY] = ".supportsExpressionsInOrderBy"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_EXTENDED_SQL_GRAMMAR] = ".supportsExtendedSQLGrammar"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_FULL_OUTER_JOINS] = ".supportsFullOuterJoins"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_GROUP_BY] = ".supportsGroupBy"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_GROUP_BY_BEYOND_SELECT] = ".supportsGroupByBeyondSelect"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_GROUP_BY_UNRELATED] = ".supportsGroupByUnrelated"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_INTEGRITY_ENHANCMENT_FACILITY] = ".supportsIntegrityEnhancementFacility"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_LIKE_ESCAPE_CLAUSE] = ".supportsLikeEscapeClause"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_LIMITED_OUTER_JOINS] = ".supportsLimitedOuterJoins"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_MINIMUM_SQL_GRAMMAR] = ".supportsMinimumSQLGrammar"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_MIXED_CASE_IDS] = ".supportsMixedCaseIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_MIXED_CASE_QUOTED_IDS] = ".supportsMixedCaseQuotedIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_MULTIPLE_RESULT_SETS] = ".supportsMultipleResultSets"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_MULTIPLE_TRANSACTIONS] = ".supportsMultipleTransactions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_NON_NULLABLE_COLS] = ".supportsNonNullableColumns"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_OPEN_CURSORS_ACROSS_COMMIT] = ".supportsOpenCursorsAcrossCommit"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_OPEN_CURSORS_ACROSS_ROLLBACK] = ".supportsOpenCursorsAcrossRollback"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_OPEN_STATEMENTS_ACROSS_COMMIT] = ".supportsOpenStatementsAcrossCommit"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_OPEN_STATEMENTS_ACROSS_ROLLBACK] = ".supportsOpenStatementsAcrossRollback"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ORDER_BY_UNRELATED] = ".supportsOrderByUnrelated"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_OUTER_JOINS] = ".supportsOuterJoins"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_POSITIONED_DELETE] = ".supportsPositionedDelete"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_POSITIONED_UPDATE] = ".supportsPositionedUpdate"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_RESULT_SET_CONCURRENCY] = ".supportsResultSetConcurrency"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_RESULT_SET_TYPE] = ".supportsResultSetType"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_DATA_MANIPULATION] = ".supportsSchemasInDataManipulation"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_INDEX_DEFINITIONS] = ".supportsSchemasInIndexDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_PRIVILEGE_DEFINITIONS] = ".supportsSchemasInPrivilegeDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_PROCEDURE_CALLS] = ".supportsSchemasInProcedureCalls"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_TABLE_DEFINITION] = ".supportsSchemasInTableDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SELECT_FOR_UPDATE] = ".supportsSelectForUpdate"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_STORED_PROCEDURES] = ".supportsStoredProcedures"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_COMPARISONS] = ".supportsSubqueriesInComparisons"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_EXISTS] = ".supportsSubqueriesInExists"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_INS] = ".supportsSubqueriesInIns"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_QUANTIFIEDS] = ".supportsSubqueriesInQuantifieds"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_TABLE_CORRELATION_NAMES] = ".supportsTableCorrelationNames"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_TRANSACTION_ISOLATION_LEVEL] = ".supportsTransactionIsolationLevel"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_TRANSACTIONS] = ".supportsTransactions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_TYPE_CONVERSION] = ".supportsTypeConversion"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_UNION] = ".supportsUnion"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_UNION_ALL] = ".supportsUnionAll"
	$asMetaQueries[$LOB_DBASE_META_UPDATES_ARE_DETECTED] = ".updatesAreDetected"
	$asMetaQueries[$LOB_DBASE_META_USES_LOCAL_FILE_PER_TABLE] = ".usesLocalFilePerTable"
	$asMetaQueries[$LOB_DBASE_META_USES_LOCAL_FILES] = ".usesLocalFiles"

	Return SetError($__LO_STATUS_SUCCESS, 0, $asMetaQueries[$iQuery])
EndFunc   ;==>__LOBase_DatabaseMetaGetQuery

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LOBase_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _LOBase_ComError_UserFunction(Default)
	Local $vUserFunction, $avUserParams[2] = ["CallArgArray", $oComError]

	If IsArray($avUserFunction) Then
		$vUserFunction = $avUserFunction[0]
		ReDim $avUserParams[UBound($avUserFunction) + 1]
		For $i = 1 To UBound($avUserFunction) - 1
			$avUserParams[$i + 1] = $avUserFunction[$i]
		Next
	Else
		$vUserFunction = $avUserFunction
	EndIf
	If IsFunc($vUserFunction) Then
		Switch $vUserFunction
			Case ConsoleWrite
				ConsoleWrite("!--COM Error-Begin--" & @CRLF & _
						"Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline & @CRLF & _
						"!--COM-Error-End--" & @CRLF)
			Case MsgBox
				MsgBox(0, "COM Error", "Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline)
			Case Else
				Call($vUserFunction, $avUserParams)
		EndSwitch
	EndIf
EndFunc   ;==>__LOBase_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_IntIsBetween
; Description ...: Test whether an input is an Integer and is between two Integers.
; Syntax ........: __LOBase_IntIsBetween($iTest, $iMin, $iMax[, $vNot = ""[, $vIncl = ""]])
; Parameters ....: $iTest               - an integer value. The Value to test.
;                  $iMin                - an integer value. The minimum $iTest can be.
;                  $iMax                - [optional] an integer value. Default is 0. The maximum $iTest can be.
;                  $vNot                - [optional] a variant value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers inside the min/max range that are not allowed.
;                  $vIncl               - [optional] a variant value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers Outside the min/max range that are allowed.
; Return values .: Success: Boolean
;                  Failure: False
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the input is between Min and Max or is an allowed number, and not one of the disallowed numbers, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_IntIsBetween($iTest, $iMin, $iMax = 0, $vNot = "", $vIncl = "")
	If Not IsInt($iTest) Then Return False

	Switch @NumParams

		Case 2
			Return ($iTest < $iMin) ? (False) : (True)

		Case 3
			Return (($iTest < $iMin) Or ($iTest > $iMax)) ? (False) : (True)

		Case 4, 5

			If IsString($vNot) Then
				If StringInStr(":" & $vNot & ":", ":" & $iTest & ":") Then Return False

			ElseIf IsInt($vNot) Then
				If ($iTest = $vNot) Then Return False

			EndIf

			If (($iTest >= $iMin) And ($iTest <= $iMax)) Then Return True

			If @NumParams = 5 Then ContinueCase

			Return False

		Case Else
			If IsString($vIncl) Then
				If StringInStr(":" & $vIncl & ":", ":" & $iTest & ":") Then Return True

			ElseIf IsInt($vIncl) Then

				If ($iTest = $vIncl) Then Return True
			EndIf

			Return False

	EndSwitch
EndFunc   ;==>__LOBase_IntIsBetween

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_SetPropertyValue
; Description ...: Creates a property value struct object.
; Syntax ........: __LOBase_SetPropertyValue($sName, $vValue)
; Parameters ....: $sName               - a string value. Property name.
;                  $vValue              - a variant value. Property value.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sName not a string
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Properties Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Property Object Returned
; Author ........: Leagnus, GMK
; Modified ......: donnyh13 - added CreateStruct function. Modified variable names.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_SetPropertyValue($sName, $vValue)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tProperties

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	$tProperties = __LOBase_CreateStruct("com.sun.star.beans.PropertyValue")
	If @error Or Not IsObj($tProperties) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$tProperties.Name = $sName
	$tProperties.Value = $vValue

	Return SetError($__LO_STATUS_SUCCESS, 0, $tProperties)
EndFunc   ;==>__LOBase_SetPropertyValue

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_VarsAreNull
; Description ...: Tests whether all input parameters are equal to Null keyword.
; Syntax ........: __LOBase_VarsAreNull($vVar1[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null[, $vVar9 = Null[, $vVar10 = Null[, $vVar11 = Null[, $vVar12 = Null]]]]]]]]]]])
; Parameters ....: $vVar1               - a variant value.
;                  $vVar2               - [optional] a variant value. Default is Null.
;                  $vVar3               - [optional] a variant value. Default is Null.
;                  $vVar4               - [optional] a variant value. Default is Null.
;                  $vVar5               - [optional] a variant value. Default is Null.
;                  $vVar6               - [optional] a variant value. Default is Null.
;                  $vVar7               - [optional] a variant value. Default is Null.
;                  $vVar8               - [optional] a variant value. Default is Null.
;                  $vVar9               - [optional] a variant value. Default is Null.
;                  $vVar10              - [optional] a variant value. Default is Null.
;                  $vVar11              - [optional] a variant value. Default is Null.
;                  $vVar12              - [optional] a variant value. Default is Null.
; Return values .: Success: Boolean
;                  Failure: False
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If All parameters are Equal to Null, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_VarsAreNull($vVar1, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null, $vVar9 = Null, $vVar10 = Null, $vVar11 = Null, $vVar12 = Null)
	Local $bAllNull1, $bAllNull2, $bAllNull3
	$bAllNull1 = (($vVar1 = Null) And ($vVar2 = Null) And ($vVar3 = Null) And ($vVar4 = Null)) ? (True) : (False)
	If (@NumParams <= 4) Then Return ($bAllNull1) ? (True) : (False)
	$bAllNull2 = (($vVar5 = Null) And ($vVar6 = Null) And ($vVar7 = Null) And ($vVar8 = Null)) ? (True) : (False)
	If (@NumParams <= 8) Then Return ($bAllNull1 And $bAllNull2) ? (True) : (False)
	$bAllNull3 = (($vVar9 = Null) And ($vVar10 = Null) And ($vVar11 = Null) And ($vVar12 = Null)) ? (True) : (False)
	Return ($bAllNull1 And $bAllNull2 And $bAllNull3) ? (True) : (False)
EndFunc   ;==>__LOBase_VarsAreNull

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_VersionCheck
; Description ...: Test if the currently installed LibreOffice version is high enough to support a certain function.
; Syntax ........: __LOBase_VersionCheck($fRequiredVersion)
; Parameters ....: $fRequiredVersion    - a floating point value. The version of LibreOffice required.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $fRequiredVersion not a Number.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Current L.O. Version.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the Current L.O. version is higher than or equal to the required version, then True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_VersionCheck($fRequiredVersion)
	Local Static $sCurrentVersion = _LOBase_VersionGet(True, False)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, False)
	Local Static $fCurrentVersion = Number($sCurrentVersion)

	If Not IsNumber($fRequiredVersion) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, False)

	Return SetError($__LO_STATUS_SUCCESS, 1, ($fCurrentVersion >= $fRequiredVersion) ? (True) : (False))
EndFunc   ;==>__LOBase_VersionCheck
