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

; Prepared Statement Input Type Commands.
Global Enum _
		$LOB_DATA_SET_TYPE_NULL, _                              ; Sets the content of the column to NULL.
		$LOB_DATA_SET_TYPE_BOOL, _                              ; Puts the given logical value into the SQL command.
		$LOB_DATA_SET_TYPE_BYTE, _                              ; Puts the given byte into the SQL command.
		$LOB_DATA_SET_TYPE_SHORT, _                             ; Puts the given integer into the SQL command.
		$LOB_DATA_SET_TYPE_INT, _                               ; Puts the given integer into the SQL command.
		$LOB_DATA_SET_TYPE_LONG, _                              ; Puts the given integer into the SQL command.
		$LOB_DATA_SET_TYPE_FLOAT, _                             ; Puts the given decimal number into the SQL command.
		$LOB_DATA_SET_TYPE_DOUBLE, _                            ; Puts the given decimal number into the SQL command.
		$LOB_DATA_SET_TYPE_STRING, _                            ; Puts the given character string into the SQL command.
		$LOB_DATA_SET_TYPE_BYTES, _                             ; Puts the given byte array into the SQL command.
		$LOB_DATA_SET_TYPE_DATE, _                              ; Puts the given date into the SQL command.
		$LOB_DATA_SET_TYPE_TIME, _                              ; Puts the given time into the SQL command.
		$LOB_DATA_SET_TYPE_TIMESTAMP, _                         ; Puts the given timestamp into the SQL command.
		$LOB_DATA_SET_TYPE_CLOB, _                              ; Puts the given CLOB (Character Large Object) into the SQL command.
		$LOB_DATA_SET_TYPE_BLOB, _                              ; Puts the given BLOB (Binary Large Object) into the SQL command.
		$LOB_DATA_SET_TYPE_ARRAY, _                             ; Puts the given Array into the SQL command.
		$LOB_DATA_SET_TYPE_OBJECT                               ; Puts the given Object into the SQL command.

; Database Data Types

Global Const _; com.sun.star.sdbc.DataType Constant Group
		$LOB_DATA_TYPE_LONGNVARCHAR = -16, _                    ; L.O. 24.2
		$LOB_DATA_TYPE_NCHAR = -15, _                           ; L.O. 24.2
		$LOB_DATA_TYPE_NVARCHAR = -9, _                         ; L.O. 24.2
		$LOB_DATA_TYPE_ROWID = -8, _                            ; L.O. 24.2
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
		$LOB_DATA_TYPE_DATALINK = 70, _                         ; L.O. 24.2
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
		$LOB_DATA_TYPE_SQLXML = 2009, _                         ; L.O. 24.2
		$LOB_DATA_TYPE_NCLOB = 2011, _                          ; L.O. 24.2
		$LOB_DATA_TYPE_REF_CURSOR = 2012, _                     ; L.O. 24.2
		$LOB_DATA_TYPE_TIME_WITH_TIMEZONE = 2013, _             ; L.O. 24.2
		$LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE = 2014           ; L.O. 24.2

; Path Convert Constants.
Global Const _
		$LOB_PATHCONV_AUTO_RETURN = 0, _                        ; Automatically returns the opposite of the input path, determined by StringInStr search for either "File:///"(L.O.Office URL) or "[A-Z]:\" (Windows File Path).
		$LOB_PATHCONV_OFFICE_RETURN = 1, _                      ; Returns L.O. Office URL, even if the input is already in that format.
		$LOB_PATHCONV_PCPATH_RETURN = 2                         ; Returns Windows File Path, even if the input is already in that format.

; Result Set Cursor Movement Commands.
Global Enum _
		$LOB_RESULT_CURSOR_MOVE_BEFORE_FIRST, _                 ; Moves before the first row.
		$LOB_RESULT_CURSOR_MOVE_FIRST, _                        ; Moves to the first row.
		$LOB_RESULT_CURSOR_MOVE_PREVIOUS, _                     ; Moves back one row.
		$LOB_RESULT_CURSOR_MOVE_NEXT, _                         ; Moves forward one row.
		$LOB_RESULT_CURSOR_MOVE_LAST, _                         ; Moves to the last record.
		$LOB_RESULT_CURSOR_MOVE_AFTER_LAST, _                   ; Moves after the last record.
		$LOB_RESULT_CURSOR_MOVE_ABSOLUTE, _                     ; Moves to the row with the given row number.
		$LOB_RESULT_CURSOR_MOVE_RELATIVE                        ; Moves backwards or forwards by the given amount: forwards for a positive value, and backwards for a negative value.

; Result Set Cursor Queries.
Global Enum _
		$LOB_RESULT_CURSOR_QUERY_IS_BEFORE_FIRST, _             ; Is the cursor before the first record. This is the case if it has not yet been reset after entry.
		$LOB_RESULT_CURSOR_QUERY_IS_FIRST, _                    ; Is the cursor on the first entry.
		$LOB_RESULT_CURSOR_QUERY_IS_LAST, _                     ; Is the cursor on the last entry.
		$LOB_RESULT_CURSOR_QUERY_IS_AFTER_LAST, _               ; Is the cursor after the last row when it is moved on with next.
		$LOB_RESULT_CURSOR_QUERY_GET_ROW                        ; Retrieve the current row number.

; Column Nullability Constants.
Global Const _; com.sun.star.sdbc.ColumnValue Constant Group
		$LOB_RESULT_METADATA_COLUMN_NOT_NULLABLE = 0, _; The column does not allow NULL values.
		$LOB_RESULT_METADATA_COLUMN_NULLABLE = 1, _; The column does allow NULL values.
		$LOB_RESULT_METADATA_COLUMN_UNKNOWN_NULLABLE = 2;  The nullability of the column is unknown.

; Column Metadata Query
Global Enum _
		$LOB_RESULT_METADATA_QUERY_GET_CATALOG_NAME, _; Gets a column's table's catalog name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_SCHEMA_NAME, _; Gets a column's table's schema. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_TABLE_NAME, _; Gets a column's table name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_DISP_SIZE, _; Gets the column's normal max width in chars. Returns an Integer.
		$LOB_RESULT_METADATA_QUERY_GET_LABEL, _; Gets the suggested column title for use in printouts and displays. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_NAME, _; Gets a column's name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_TYPE, _; Gets the column's SQL type. Returns an Integer. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
		$LOB_RESULT_METADATA_QUERY_GET_TYPE_NAME, _; Gets the column's database-specific type name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_LENGTH, _; Gets a column's number of decimal digits. Returns an Integer.
		$LOB_RESULT_METADATA_QUERY_GET_DECIMAL_PLACE, _; Gets a column's number of digits to right of the decimal point. Returns an Integer.
		$LOB_RESULT_METADATA_QUERY_IS_AUTO_VALUE, _; Query whether the column is automatically numbered, thus read-only (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_CASE_SENSITIVE, _; Query whether a column's case matters (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_CURRENCY, _; Query whether the column is a cash value (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_NULLABLE, _; Query the nullability of values in the designated column. Returns an Integer. See Constants, $LOB_RESULT_METADATA_COLUMN_* as defined in LibreOfficeBase_Constants.au3.
		$LOB_RESULT_METADATA_QUERY_IS_READ_ONLY, _; Query whether a column is definitely not writable (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_WRITABLE, _; Query whether it is possible for a write on the column to succeed (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_WRITABLE_DEFINITE, _; Query whether a write on the column will definitely succeed (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_SEARCHABLE, _; Query whether the column can be used in a where clause (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_SIGNED; Query whether values in the column are signed numbers (True if so). Returns a Boolean.

; Result Set Row Modification Commands.
Global Enum _
		$LOB_RESULT_ROW_MOD_NULL, _                             ; Sets the column content to NULL
		$LOB_RESULT_ROW_MOD_BOOL, _                             ; Changes the content of specified column to the called logical value.
		$LOB_RESULT_ROW_MOD_BYTE, _                             ; Changes the content of specified column to the called byte.
		$LOB_RESULT_ROW_MOD_SHORT, _                            ; Changes the content of specified column to the called integer.
		$LOB_RESULT_ROW_MOD_INT, _                              ; Changes the content of specified column to the called integer.
		$LOB_RESULT_ROW_MOD_LONG, _                             ; Changes the content of specified column to the called integer.
		$LOB_RESULT_ROW_MOD_FLOAT, _                            ; Changes the content of specified column to the called decimal number.
		$LOB_RESULT_ROW_MOD_DOUBLE, _                           ; Changes the content of specified column to the called decimal number.
		$LOB_RESULT_ROW_MOD_STRING, _                           ; Changes the content of specified column to the called string.
		$LOB_RESULT_ROW_MOD_BYTES, _                            ; Changes the content of specified column to the called byte array.
		$LOB_RESULT_ROW_MOD_DATE, _                             ; Changes the content of specified column to the called date.
		$LOB_RESULT_ROW_MOD_TIME, _                             ; Changes the content of specified column to the called time.
		$LOB_RESULT_ROW_MOD_TIMESTAMP                           ; Changes the content of specified column to the called Date and Time (Timestamp).

; Result Set Queries.
Global Enum _
		$LOB_RESULT_ROW_QUERY_IS_ROW_INSERTED, _                ; Indicates if this is a new row.
		$LOB_RESULT_ROW_QUERY_IS_ROW_UPDATED, _                 ; Indicates if the current row has been altered.
		$LOB_RESULT_ROW_QUERY_IS_ROW_DELETED                    ; Indicates if the current row has been deleted.

; Result Set Row Read Commands.
Global Enum _
		$LOB_RESULT_ROW_READ_STRING, _                          ; Returns the content of the column as a character string.
		$LOB_RESULT_ROW_READ_BOOL, _                            ; Returns the content of the column as a boolean value.
		$LOB_RESULT_ROW_READ_BYTE, _                            ; Returns the content of the column as a single byte.
		$LOB_RESULT_ROW_READ_SHORT, _                           ; Returns the content of the column as an integer.
		$LOB_RESULT_ROW_READ_INT, _                             ; Returns the content of the column as an integer.
		$LOB_RESULT_ROW_READ_LONG, _                            ; Returns the content of the column as an integer.
		$LOB_RESULT_ROW_READ_FLOAT, _                           ; Returns the content of the column as a single precision decimal number.
		$LOB_RESULT_ROW_READ_DOUBLE, _                          ; Returns the content of the column as a double precision decimal number.
		$LOB_RESULT_ROW_READ_BYTES, _                           ; Returns the content of the column as an array of single bytes.
		$LOB_RESULT_ROW_READ_DATE, _                            ; Returns the content of the column as a date.
		$LOB_RESULT_ROW_READ_TIME, _                            ; Returns the content of the column as a time value.
		$LOB_RESULT_ROW_READ_TIMESTAMP, _                       ; Returns the content of the column as a timestamp (date and time).
		$LOB_RESULT_ROW_READ_WAS_NULL                           ; Indicates if the value of the most recently read column was NULL.

; Result Set Row Update Commands.
Global Enum _
		$LOB_RESULT_ROW_UPDATE_INSERT, _                        ; Saves a new row.
		$LOB_RESULT_ROW_UPDATE_UPDATE, _                        ; Confirms alteration of the current row.
		$LOB_RESULT_ROW_UPDATE_DELETE, _                        ; Deletes the current row.
		$LOB_RESULT_ROW_UPDATE_CANCEL_UPDATE, _                 ; Reverses changes in the current row.
		$LOB_RESULT_ROW_UPDATE_MOVE_TO_INSERT, _                ; Moves the cursor into a row corresponding to a new record.
		$LOB_RESULT_ROW_UPDATE_MOVE_TO_CURRENT                  ; After the entry of a new record, returns the cursor to its previous position.

