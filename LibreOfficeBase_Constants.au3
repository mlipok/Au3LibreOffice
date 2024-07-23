#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#include-once

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Base Constants for the Libre Office UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various functions in the Libre Office UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; Sleep Divisor $__LOCCONST_SLEEP_DIV
; In applicable functions this is used for adjusting how frequent a sleep occurs in loops.
; For any number above 0 the number of times a loop has completed is divided by $__LOBCONST_SLEEP_DIV. If you find some functions cause momentary freeze ups, a recommended value is 15.
; Set to 0 for no pause in a loop.
Global Const $__LOBCONST_SLEEP_DIV = 0

#Tidy_ILC_Pos=65

; Database Data Types
Global Const _
		$LOB_DATA_TYPE_LONGNVARCHAR = -16, _ ; L.O. 24.2
		$LOB_DATA_TYPE_NCHAR = -15, _ ; L.O. 24.2
		$LOB_DATA_TYPE_NVARCHAR = -9, _ ; L.O. 24.2
		$LOB_DATA_TYPE_ROWID = -8, _ ; L.O. 24.2
		$LOB_DATA_TYPE_BIT = -7, _
		$LOB_DATA_TYPE_TINYINT = -6, _
		$LOB_DATA_TYPE_BIGINT = -5, _
		$LOB_DATA_TYPE_LONGVARBINARY = -4, _
		$LOB_DATA_TYPE_VARBINARY = -3, _
		$LOB_DATA_TYPE_BINARY = -2, _
		$LOB_DATA_TYPE_LONGVARCHAR = -1, _
		$LOB_DATA_TYPE_SQLNULL = 0, _
		$LOB_DATA_TYPE_CHAR = 1, _
		$LOB_DATA_TYPE_NUMERIC = 2, _
		$LOB_DATA_TYPE_DECIMAL = 3, _
		$LOB_DATA_TYPE_INTEGER = 4, _
		$LOB_DATA_TYPE_SMALLINT = 5, _
		$LOB_DATA_TYPE_FLOAT = 6, _
		$LOB_DATA_TYPE_REAL = 7, _
		$LOB_DATA_TYPE_DOUBLE = 8, _
		$LOB_DATA_TYPE_VARCHAR = 12, _
		$LOB_DATA_TYPE_BOOLEAN = 16, _
		$LOB_DATA_TYPE_DATALINK = 70, _ ; L.O. 24.2
		$LOB_DATA_TYPE_DATE = 91, _
		$LOB_DATA_TYPE_TIME = 92, _
		$LOB_DATA_TYPE_TIMESTAMP = 93, _
		$LOB_DATA_TYPE_OTHER = 1111, _
		$LOB_DATA_TYPE_OBJECT = 2000, _
		$LOB_DATA_TYPE_DISTINCT = 2001, _
		$LOB_DATA_TYPE_STRUCT = 2002, _
		$LOB_DATA_TYPE_ARRAY = 2003, _
		$LOB_DATA_TYPE_BLOB = 2004, _
		$LOB_DATA_TYPE_CLOB = 2005, _
		$LOB_DATA_TYPE_REF = 2006, _
		$LOB_DATA_TYPE_SQLXML = 2009, _ ; L.O. 24.2
		$LOB_DATA_TYPE_NCLOB = 2011, _ ; L.O. 24.2
		$LOB_DATA_TYPE_REF_CURSOR = 2012, _ ; L.O. 24.2
		$LOB_DATA_TYPE_TIME_WITH_TIMEZONE = 2013, _ ; L.O. 24.2
		$LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE = 2014 ; L.O. 24.2

; Path Convert Constants.
Global Const _
		$LOB_PATHCONV_AUTO_RETURN = 0, _                        ; Automatically returns the opposite of the input path, determined by StringInStr search for either "File:///"(L.O.Office URL) or "[A-Z]:\" (Windows File Path).
		$LOB_PATHCONV_OFFICE_RETURN = 1, _                      ; Returns L.O. Office URL, even if the input is already in that format.
		$LOB_PATHCONV_PCPATH_RETURN = 2                         ; Returns Windows File Path, even if the input is already in that format.

