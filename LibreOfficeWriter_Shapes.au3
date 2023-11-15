#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Internal.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Inserting shapes in L.O. Writer.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_ShapesGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapesGetNames
; Description ...: Return a list of Shape names contained in a document.
; Syntax ........: _LOWriter_ShapesGetNames(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object.  A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 2D Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Shapes Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning 2D Array containing a list of Shape names contained in a document, the first column ($aArray[0][0] contains the shape name, the second column ($aArray[0][1] contains the shape's Implementation name. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Implementation name identifies what type of shape object it is, as there can be multiple things counted as "Shapes", such as Text Frames etc.
;				   I have found the three Implementation names being returned, SwXTextFrame, indicating the shape is actually a Text Frame, SwXShape, is a regular shape such as a line, circle etc. And "SwXTextGraphicObject", which is an image / picture. There may be other return types I haven't found yet.
;				   Images inserted into the document are also listed as TextFrames in the shapes category. There isn't an easy way to differentiate between them yet, see _LOWriter_FramesGetNames, to search for Frames in the shapes category.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapesGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asShapeNames[0][2]
	Local $oShapes

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oShapes = $oDoc.DrawPage()
	If Not IsObj($oShapes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $oShapes.hasElements() Then
		ReDim $asShapeNames[$oShapes.getCount()][2]
		For $i = 0 To $oShapes.getCount() - 1
			$asShapeNames[$i][0] = $oShapes.getByIndex($i).Name()
			If $oShapes.getByIndex($i).supportsService("com.sun.star.drawing.Text") Then
				; If Supports Text Method, then get that impl. name, else just the regular impl. name.
				$asShapeNames[$i][1] = $oShapes.getByIndex($i).Text.ImplementationName()
			Else
				$asShapeNames[$i][1] = $oShapes.getByIndex($i).ImplementationName()
			EndIf

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, UBound($asShapeNames), $asShapeNames)
EndFunc   ;==>_LOWriter_ShapesGetNames



; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapePoints
; Description ...: Set or Retrieve a Shape's Position Points.
; Syntax ........: _LOWriter_ShapePoints(ByRef $oShape[, $avPoints = Null])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function. See Remarks.
;                  $avPoints            - [optional] an array of variants. Default is Null. A two column Array of Position Points and Point Type Constants, previously returned from this function. Call with Null to retrieve the current Position Point Array.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oShape does not have property "PolyPolygonBezier", and consequently does not have Position Points.
;				   @Error 1 @Extended 3 Return 0 = $avPoints not an Array, and not set to Null.
;				   @Error 1 @Extended 4 Return 0 = $avPoints Array does not have two columns.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Array of Position Points from Shape.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Array of Point Type Constants from Shape.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve PolyPolygonBezier Structure from Shape.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Position Points were successfully set.
;				   @Error 0 @Extended ? Return Array = Success. $avPoints was set to Null, returning current Position Points and corresponding Point Type Constants in a 2 Column Array with Positions listed from first to last. @Extended is set to the number of rows contained in the Array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Position Point determines where the Shape's lines are drawn. A Position point is a combination of an X and Y position, and a Point Type Constant indicating what type of Point it is.
;				   This function returns a two column Array, the first column contains the Position Point Structure with X and Y properties, and the second column has the corresponding Point Type Constant.
;				   Only $LOW_SHAPE_TYPE_LINE_* type shapes have Points that can be modified, and the Shape Line type $LOW_SHAPE_TYPE_LINE_LINE, can only have two points, the beginning and the end.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName, _LOWriter_ShapePointsAdd, _LOWriter_ShapePointsRemove, _LOWriter_ShapePointsModify
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapePoints(ByRef $oShape, $avPoints = Null)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $UBOUND_COLUMNS = 2
	Local $tPolyCoords
	Local $atPoints[0]
	Local $aiFlags[0]
	Local $avArray[1]

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If Not ($oShape.getPropertySetInfo().hasPropertyByName("PolyPolygonBezier")) Then Return SetError($__LOW_STATUS_INPUT_ERROR,2,0)

	If ($avPoints = Null) Then
		; Retrieve the Array of Position Points.
	$atPoints = $oShape.PolyPolygonBezier.Coordinates()[0]
If Not IsArray($atPoints) Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)
; Retrieve the Array of Point Type Constants.
$aiFlags =  $oShape.PolyPolygonBezier.Flags()[0]
If Not IsArray($atPoints) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

ReDim $avArray[UBound($atPoints)][2]; Convert $avArray to two columns.

For $i = 0 To UBound($avArray) -1
; Fill the Array with Position Points and corresponding Point Types.
$avArray[$i][0] = $atPoints[$i]
$avArray[$i][1] = $aiFlags[$i]

Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
Next

Return SetError($__LOW_STATUS_SUCCESS,UBound($avArray),$avArray)

		EndIf

If Not IsArray($avPoints) Then Return SetError($__LOW_STATUS_INPUT_ERROR,3,0)
If (UBound($avPoints,$UBOUND_COLUMNS) <> 2) Then Return SetError($__LOW_STATUS_INPUT_ERROR,4,0)

ReDim $atPoints[UBound($avPoints)]
ReDim $aiFlags[UBound($avPoints)]

For $i = 0 To UBound($avPoints) - 1
; Fill the Individual Points and Flags Arrays with values contained in $avPoints.
$atPoints[$i] = $avPoints[$i][0]
$aiFlags[$i] = $avPoints[$i][1]

Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
Next

$tPolyCoords = $oShape.PolyPolygonBezier()
If Not IsObj($tPolyCoords) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

; Each Array needs to be nested in an array.
$avArray[0] = $atPoints
$tPolyCoords.Coordinates = $avArray

$avArray[0] = $aiFlags
$tPolyCoords.Flags = $avArray

; Set the  new Position Points for the Shape.
$oShape.PolyPolygonBezier = $tPolyCoords

Return SetError($__LOW_STATUS_SUCCESS,0,1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapePointsAdd
; Description ...: Add a Position Point to an Array of Position Points.
; Syntax ........: _LOWriter_ShapePointsAdd(ByRef $avPoints, $iArrayElement, $iX, $iY, $iPointType)
; Parameters ....: $avPoints            - [in/out] an array of variants. A two column Array of Position Points and Point Types returned from _LOWriter_ShapePoints. Array will be directly modified.
;                  $iArrayElement       - an integer value. The Array Element to add the new point before. See Remarks.
;                  $iX                  - an integer value. The X coordinate value, set in Micrometers.
;                  $iY                  - an integer value. The Y coordinate value, set in Micrometers.
;                  $iPointType          - an integer value (0-3). The Type of Point this new Point is. See Remarks. See constants $LOW_SHAPE_POINT_TYPE_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $avPoints not an Array.
;				   @Error 1 @Extended 2 Return 0 = Array called in $avPoints does not contain two columns.
;				   @Error 1 @Extended 3 Return 0 = $iArrayElement is not an Integer, less than 0, or greater than number of elements contained in the Array plus 1.
;				   @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iY not an Integer
;				   @Error 1 @Extended 6 Return 0 = $PointType not an Integer, less than 0 or greater than 3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create a new Position Point Structure.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. New Position Point was successfully added to the Array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iArrayElement is the Array Element to add the new Position Point before, i.e. to add a new point at the beginning of the Array, you would call $iArrayElement with 0, to add a new point to the end, you would call $iArrayElement with the last element number of the Array plus 1.
;				   According to Andrew Pitonyak (OOME. 3.0, page 584), not all Point Type Constants (flags) combinations are valid for Position values.
;					"Not all combinations of points and flags are valid. A complete discussion of what constitutes a valid combination of points and flags is beyond the scope of this book."
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapePointsAdd(ByRef $avPoints, $iArrayElement, $iX, $iY, $iPointType)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $UBOUND_COLUMNS = 2
	Local $tPoint
	Local $iOffset = 0
	Local $avArray[0][2]

If Not IsArray($avPoints) Then Return SetError($__LOW_STATUS_INPUT_ERROR,1,0)
If (UBound($avPoints,$UBOUND_COLUMNS) <> 2) Then Return SetError($__LOW_STATUS_INPUT_ERROR,2,0)
If Not __LOWriter_IntIsBetween($iArrayElement,0,UBound($avPoints)) Then  Return SetError($__LOW_STATUS_INPUT_ERROR,3,0)
If Not IsInt($iX) Then  Return SetError($__LOW_STATUS_INPUT_ERROR,4,0)
If Not IsInt($iY) Then  Return SetError($__LOW_STATUS_INPUT_ERROR,5,0)
If Not __LOWriter_IntIsBetween($iPointType,$LOW_SHAPE_POINT_TYPE_NORMAL,$LOW_SHAPE_POINT_TYPE_SYMMETRIC) Then  Return SetError($__LOW_STATUS_INPUT_ERROR,6,0)

$tPoint = __LOWriter_CreatePoint($iX, $iY)
If @error Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

ReDim $avArray[UBound($avPoints) + 1][2]

For $i = 0 To UBound($avArray) -1

If ($i = $iArrayElement) Then
$avArray[$i][0] = $tPoint
$avArray[$i][1] = $iPointType
$iOffset -= 1

Else
$avArray[$i][0] = $avPoints[$i + $iOffset][0]
$avArray[$i][1] = $avPoints[$i + $iOffset][1]

	EndIf

Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
Next

$avPoints = $avArray

Return SetError($__LOW_STATUS_SUCCESS,0,1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapePointsRemove
; Description ...: Remove a position Point from an Array of Position Points.
; Syntax ........: _LOWriter_ShapePointsRemove(ByRef $avPoints, $iArrayElement)
; Parameters ....: $avPoints            - [in/out] an array of variants. A two column Array of Position Points and Point Types returned from _LOWriter_ShapePoints. Array will be directly modified.
;                  $iArrayElement       - an integer value. The Array Element to remove.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $avPoints not an Array.
;				   @Error 1 @Extended 2 Return 0 = Array called in $avPoints does not contain two columns.
;				   @Error 1 @Extended 3 Return 0 = $iArrayElement is not an Integer, less than 0, or greater than number of elements contained in the Array.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Position Point was successfully deleted from the Array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapePointsRemove(ByRef $avPoints, $iArrayElement)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $UBOUND_COLUMNS = 2
	Local $iOffset = 0
	Local $avArray[0][2]

If Not IsArray($avPoints) Then Return SetError($__LOW_STATUS_INPUT_ERROR,1,0)
If (UBound($avPoints,$UBOUND_COLUMNS) <> 2) Then Return SetError($__LOW_STATUS_INPUT_ERROR,2,0)
If Not __LOWriter_IntIsBetween($iArrayElement,0,(UBound($avPoints) - 1)) Then  Return SetError($__LOW_STATUS_INPUT_ERROR,3,0)

ReDim $avArray[UBound($avPoints) - 1][2]

For $i = 0 To UBound($avPoints) -1

If ($i = $iArrayElement) Then ; Skip that element
$iOffset -= 1

Else
$avArray[$i + $iOffset][0] = $avPoints[$i][0]
$avArray[$i + $iOffset][1] = $avPoints[$i][1]

	EndIf

Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
Next

$avPoints = $avArray

Return SetError($__LOW_STATUS_SUCCESS,0,1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapePointsModify
; Description ...: Modify an existing Position Point or Point Type in an Array of Position Points.
; Syntax ........: _LOWriter_ShapePointsModify(ByRef $avPoints, $iArrayElement[, $iX = Null[, $iY = Null[, $iPointType = Null]]])
; Parameters ....: $avPoints            - [in/out] an array of variants. A two column Array of Position Points and Point Types returned from _LOWriter_ShapePoints. Array will be directly modified.
;                  $iArrayElement       - an integer value. The Array Element to modify the point of.
;                  $iX                  - [optional] an integer value. Default is Null. The X coordinate value, set in Micrometers.
;                  $iY                  - [optional] an integer value. Default is Null. The Y coordinate value, set in Micrometers.
;                  $iPointType          - [optional] an integer value. Default is Null. The Type of Point this new Point is. See Remarks. See constants $LOW_SHAPE_POINT_TYPE_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $avPoints not an Array.
;				   @Error 1 @Extended 2 Return 0 = Array called in $avPoints does not contain two columns.
;				   @Error 1 @Extended 3 Return 0 = $iArrayElement is not an Integer, less than 0, or greater than number of elements contained in the Array.
;				   @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iY not an Integer
;				   @Error 1 @Extended 6 Return 0 = $PointType not an Integer, less than 0 or greater than 3.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in  the following order: the current X value, the current Y value, and the corresponding Point Type Constant.
;~
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings for the Array Element called in $iArrayElement.
;				   Call any optional parameter with Null keyword to skip it.
;				   According to Andrew Pitonyak (OOME. 3.0, page 584), not all Point Type Constants (flags) combinations are valid for Position values.
;					"Not all combinations of points and flags are valid. A complete discussion of what constitutes a valid combination of points and flags is beyond the scope of this book."
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapePointsModify(ByRef $avPoints, $iArrayElement, $iX = Null, $iY = Null, $iPointType = Null)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $UBOUND_COLUMNS = 2
	Local $avPointVals[3]

If Not IsArray($avPoints) Then Return SetError($__LOW_STATUS_INPUT_ERROR,1,0)
If (UBound($avPoints,$UBOUND_COLUMNS) <> 2) Then Return SetError($__LOW_STATUS_INPUT_ERROR,2,0)
If Not __LOWriter_IntIsBetween($iArrayElement,0,(UBound($avPoints) - 1)) Then  Return SetError($__LOW_STATUS_INPUT_ERROR,3,0)

If __LOWriter_VarsAreNull($iX, $iY, $iPointType) Then
__LOWriter_ArrayFill($avPointVals, $avPoints[$iArrayElement][0].X(), $avPoints[$iArrayElement][0].Y(), $avPoints[$iArrayElement][1])

Return SetError($__LOW_STATUS_SUCCESS,1,$avPointVals)
	EndIf

If ($iX <> Null) Then
	If Not IsInt($iX) Then  Return SetError($__LOW_STATUS_INPUT_ERROR,4,0)
	$avPoints[$iArrayElement][0].X = $iX
	EndIf

If ($iY <> Null) Then
	If Not IsInt($iY) Then  Return SetError($__LOW_STATUS_INPUT_ERROR,5,0)
	$avPoints[$iArrayElement][0].Y = $iY
	EndIf

If ($iPointType <> Null) Then
	If Not __LOWriter_IntIsBetween($iPointType,$LOW_SHAPE_POINT_TYPE_NORMAL,$LOW_SHAPE_POINT_TYPE_SYMMETRIC) Then  Return SetError($__LOW_STATUS_INPUT_ERROR,6,0)
	$avPoints[$iArrayElement][1] = $iPointType
	EndIf

Return SetError($__LOW_STATUS_SUCCESS,0,1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeLineProperties
; Description ...: Set or Retrieve Shape Line settings.
; Syntax ........: _LOWriter_ShapeLineProperties(ByRef $oShape[, $vStyle = Null[, $iColor = Null[, $iWidth = Null[, $iTransparency = Null[, $iCornerStyle = Null[, $iCapStyle = Null]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $vStyle              - [optional] a variant value (0-31, or String). Default is Null. The Line Style to use. Can be a Custom Line Style name, or one of the constants, $LOW_SHAPE_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The Line color, set in Long integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iWidth              - [optional] an integer value (0-5004). Default is Null. The line Width, set in Micrometers.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The Line transparency percentage. 100% = fully transparent.
;                  $iCornerStyle        - [optional] an integer value (0,2-4). Default is Null. The Line Corner Style. See Constants $LOW_SHAPE_LINE_JOINT_* as defined in LibreOfficeWriter_Constants.au3
;                  $iCapStyle           - [optional] an integer value (0-2). Default is Null. The Line Cap Style. See Constants $LOW_SHAPE_LINE_CAP_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $vStyle not a String, and not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $vStyle is an Integer, but less than 0, or greater than 31. See constants $LOW_SHAPE_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $iColor not an Integer, less than 0, or greater than 16777215.
;				   @Error 1 @Extended 5 Return 0 = $iWidth not an Integer, less than 0, or greater than 5004.
;				   @Error 1 @Extended 6 Return 0 = $iTransparency not an Integer, less than 0, or greater than 100.
;				   @Error 1 @Extended 7 Return 0 = $iCornerStyle not an Integer, not equal to 0, equal to 1, not equal to 2 or greater than 4.  See Constants $LOW_SHAPE_LINE_JOINT_* as defined in LibreOfficeWriter_Constants.au3
;				   @Error 1 @Extended 8 Return 0 = $iCapStyle is an Integer, but less than 0, or greater than 2. See constants $LOW_SHAPE_LINE_CAP_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to convert Constant to Line Style name.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $vStyle
;				   |								2 = Error setting $iColor
;				   |								4 = Error setting $iWidth
;				   |								8 = Error setting $iTransparency
;				   |								16 = Error setting $iCornerStyle
;				   |								32 = Error setting $iCapStyle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $vStyle accepts a String or an Integer because there is the possibility of a custom Line Style being available that the user may want to use.
;					When retrieving the current settings, $vStyle could be either an integer or a String. It will be a String if the current Line Style is a custom Line Style, else an Integer, corresponding to one of the constants, $LOW_SHAPE_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......:  _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName, _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapeLineProperties(ByRef $oShape, $vStyle = Null, $iColor = Null, $iWidth = Null, $iTransparency = Null, $iCornerStyle = Null, $iCapStyle = Null)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local Const $__LOW_SHAPE_LINE_STYLE_NONE = 0, $__LOW_SHAPE_LINE_STYLE_SOLID = 1, $__LOW_SHAPE_LINE_STYLE_DASH = 2
	Local $avLine[6]
	Local $sStyle
	Local $vReturn

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($vStyle, $iColor, $iWidth, $iTransparency, $iCornerStyle, $iCapStyle) Then

		Switch $oShape.LineStyle()

			Case $__LOW_SHAPE_LINE_STYLE_NONE

			$vReturn = $LOW_SHAPE_LINE_STYLE_NONE

			Case $__LOW_SHAPE_LINE_STYLE_SOLID

			$vReturn = $LOW_SHAPE_LINE_STYLE_CONTINUOUS

			Case $__LOW_SHAPE_LINE_STYLE_DASH

				$vReturn = __LOWriter_ShapeLineStyleName(Null, $oShape.LineDashName())
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)
				EndSwitch

			__LOWriter_ArrayFill($avLine, $vReturn, $oShape.LineColor(), $oShape.LineWidth(), $oShape.LineTransparence(), $oShape.LineJoint(), $oShape.LineCap())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avLine)
	EndIf

	If ($vStyle <> Null) Then
		If Not IsString($vStyle) And Not IsInt($vStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

		If IsInt($vStyle) Then
		If Not __LOWriter_IntIsBetween($vStyle,$LOW_SHAPE_LINE_STYLE_NONE,$LOW_SHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

		Switch $vStyle

			Case $LOW_SHAPE_LINE_STYLE_NONE

			$oShape.LineStyle = $__LOW_SHAPE_LINE_STYLE_NONE
			$iError = ($oShape.LineStyle() = $__LOW_SHAPE_LINE_STYLE_NONE) ? $iError : BitOR($iError, 1)

			Case $LOW_SHAPE_LINE_STYLE_CONTINUOUS

			$oShape.LineStyle = $__LOW_SHAPE_LINE_STYLE_SOLID
			$iError = ($oShape.LineStyle() = $__LOW_SHAPE_LINE_STYLE_SOLID) ? $iError : BitOR($iError, 1)

			Case Else

			$sStyle = __LOWriter_ShapeLineStyleName($vStyle)
			If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)
			$oShape.LineStyle = $__LOW_SHAPE_LINE_STYLE_DASH
			$oShape.LineDashName = $sStyle
			$iError = ($oShape.LineDashName() = $sStyle) ? $iError : BitOR($iError, 1)
				EndSwitch

		Else

			$sStyle = $vStyle
		$oShape.LineDashName = $sStyle
		$iError = ($oShape.LineDashName() = $sStyle) ? $iError : BitOR($iError, 1)

			EndIf

	EndIf

	If ($iColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iColor,$LOW_COLOR_BLACK,$LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oShape.LineColor = $iColor
		$iError = ($oShape.LineColor() = $iColor) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iWidth <> Null) Then
		If Not __LOWriter_IntIsBetween($iWidth,0,5004) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oShape.LineWidth = $iWidth
		$iError = (__LOWriter_IntIsBetween($oShape.LineWidth(), $iWidth - 1, $iWidth + 1)) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iTransparency <> Null) Then
		If Not __LOWriter_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	$oShape.LineTransparence = $iTransparency
		$iError = ($oShape.LineTransparence() = $iTransparency) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iCornerStyle <> Null) Then
			If Not __LOWriter_IntIsBetween($iCornerStyle,$LOW_SHAPE_LINE_JOINT_NONE, $LOW_SHAPE_LINE_JOINT_ROUND, $LOW_SHAPE_LINE_JOINT_MIDDLE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oShape.LineJoint = $iCornerStyle
		$iError = ($oShape.LineJoint() = $iCornerStyle) ? $iError : BitOR($iError, 16)
	EndIf

	If ($iCapStyle <> Null) Then
		If Not __LOWriter_IntIsBetween($iCapStyle,$LOW_SHAPE_LINE_CAP_FLAT,$LOW_SHAPE_LINE_CAP_SQUARE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oShape.LineCap = $iCapStyle
		$iError = ($oShape.LineCap() = $iCapStyle) ? $iError : BitOR($iError, 32)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeLineArrowStyles
; Description ...: Set or Retrieve Shape Line Start and End Arrow Style settings.
; Syntax ........: _LOWriter_ShapeLineArrowStyles(ByRef $oShape[, $vStartStyle = Null[, $iStartWidth = Null[, $bStartCenter = Null[, $bSync = Null[, $vEndStyle = Null[, $iEndWidth = Null[, $bEndCenter = Null]]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $vStartStyle         - [optional] a variant value (0-32, or String). Default is Null. The Arrow head to apply to the start of the line. Can be a Custom Arrowhead name, or one of the constants, $LOW_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
;                  $iStartWidth         - [optional] an integer value (0-5004). Default is Null. The Width of the Starting Arrowhead, in Micrometers.
;                  $bStartCenter        - [optional] a boolean value. Default is Null. If True, Places the center of the Start arrowhead on the endpoint of the line.
;                  $bSync               - [optional] a boolean value. Default is Null. If True, Synchronizes the Start Arrowhead settings with the end Arrowhead settings. See remarks.
;                  $vEndStyle           - [optional] a variant value (0-32, or String). Default is Null. The Arrow head to apply to the end of the line. Can be a Custom Arrowhead name, or one of the constants, $LOW_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
;                  $iEndWidth           - [optional] an integer value (0-5004). Default is Null. The Width of the Ending Arrowhead, in Micrometers.
;                  $bEndCenter          - [optional] a boolean value. Default is Null. If True, Places the center of the End arrowhead on the endpoint of the line.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $vStartStyle not a String, and not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $vStartStyle is an Integer, but less than 0, or greater than 32. See constants $LOW_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $iStartWidth not an Integer, less than 0, or greater than 5004.
;				   @Error 1 @Extended 5 Return 0 = $bStartCenter not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $bSync not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $vEndStyle not a String, and not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $vSEndStyle is an Integer, but less than 0, or greater than 32. See constants $LOW_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 9 Return 0 = $iEndWidth not an Integer, less than 0, or greater than 5004.
;				   @Error 1 @Extended 8 Return 0 = $bEndCenter not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to convert Constant to Arrowhead name.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $vStartStyle
;				   |								2 = Error setting $iStartWidth
;				   |								4 = Error setting $bStartCenter
;				   |								8 = Error setting $bSync
;				   |								16 = Error setting $vEndStyle
;				   |								32 = Error setting $iEndWidth
;				   |								64 = Error setting $bEndCenter
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Libre Office has no setting for $bSync, so I have made a manual version of it in this function. It only accepts True, and must be called with True each time you want it to synchronize.
;					When retrieving the current settings, $bSync will be a Boolean value of whether the Start Arrowhead settings are currently equal to the End Arrowhead setting values.
;				   Both $vStartStyle and $vEndStyle accept a String or an Integer because there is the possibility of a custom Arrowhead being available the user may want to use.
;					When retrieving the current settings, both $vStartStyle and $vEndStyle could be either an integer or a String. It will be a String if the current Arrowhead is a custom Arrowhead, else an Integer, corresponding to one of the constants, $LOW_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapeLineArrowStyles(ByRef $oShape, $vStartStyle = Null, $iStartWidth = Null, $bStartCenter = Null, $bSync = Null, $vEndStyle = Null, $iEndWidth = Null, $bEndCenter = Null)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avArrow[7]
	Local $sStartStyle, $sEndStyle

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($vStartStyle, $iStartWidth, $bStartCenter, $bSync, $vEndStyle, $iEndWidth, $bEndCenter) Then
			__LOWriter_ArrayFill($avArrow, __LOWriter_ShapeArrowStyleName(Null,$oShape.LineStartName()), $oShape.LineStartWidth(), $oShape.LineStartCenter(), _
				((($oShape.LineStartName() = $oShape.LineEndName()) And ($oShape.LineStartWidth() = $oShape.LineEndWidth()) And ($oShape.LineStartCenter() = $oShape.LineEndCenter())) ? True : False), _ ; See if Start and End are the same.
				__LOWriter_ShapeArrowStyleName(Null,$oShape.LineEndName()), $oShape.LineEndWidth(), $oShape.LineEndCenter())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avArrow)
	EndIf

	If ($vStartStyle <> Null) Then
		If Not IsString($vStartStyle) And Not IsInt($vStartStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		If IsInt($vStartStyle) Then
			If Not __LOWriter_IntIsBetween($vStartStyle,$LOW_SHAPE_LINE_ARROW_TYPE_NONE,$LOW_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			$sStartStyle = __LOWriter_ShapeArrowStyleName($vStartStyle)
			If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)
		Else
			$sStartStyle = $vStartStyle
			EndIf

		$oShape.LineStartName = $sStartStyle
		$iError = ($oShape.LineStartName() = $sStartStyle) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iStartWidth <> Null) Then
		If Not __LOWriter_IntIsBetween($iStartWidth,0,5004) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oShape.LineStartWidth = $iStartWidth
		$iError = (__LOWriter_IntIsBetween($oShape.LineStartWidth(), $iStartWidth -1, $iStartWidth + 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bStartCenter <> Null) Then
		If Not IsBool($bStartCenter) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oShape.LineStartCenter = $bStartCenter
		$iError = ($oShape.LineStartCenter() = $bStartCenter) ? $iError : BitOR($iError, 4)
	EndIf

	If ($bSync <> Null) Then
		If Not IsBool($bSync) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If ($bSync = True) Then
	$oShape.LineEndName = $oShape.LineStartName()
	$oShape.LineEndWidth = $oShape.LineStartWidth()
	$oShape.LineEndCenter = $oShape.LineStartCenter()
		$iError = (($oShape.LineStartName() = $oShape.LineEndName()) And _
				($oShape.LineStartWidth() = $oShape.LineEndWidth()) And _
				($oShape.LineStartCenter() = $oShape.LineEndCenter())) ? $iError : BitOR($iError, 8)
		EndIf

	EndIf

	If ($vEndStyle <> Null) Then
			If Not IsString($vEndStyle) And Not IsInt($vEndStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		If IsInt($vEndStyle) Then
			If Not __LOWriter_IntIsBetween($vEndStyle,$LOW_SHAPE_LINE_ARROW_TYPE_NONE,$LOW_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
						$sEndStyle = __LOWriter_ShapeArrowStyleName($vEndStyle)
			If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)
		Else
			$sEndStyle = $vEndStyle
			EndIf

		$oShape.LineEndName = $sEndStyle
		$iError = ($oShape.LineEndName() = $sEndStyle) ? $iError : BitOR($iError, 16)
	EndIf

	If ($iEndWidth <> Null) Then
		If Not __LOWriter_IntIsBetween($iEndWidth,0,5004) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$oShape.LineEndWidth = $iEndWidth
		$iError = (__LOWriter_IntIsBetween($oShape.LineEndWidth(), $iEndWidth -1, $iEndWidth + 1)) ? $iError : BitOR($iError, 32)
	EndIf

	If ($bEndCenter <> Null) Then
		If Not IsBool($bEndCenter) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
		$oShape.LineEndCenter = $bEndCenter
		$iError = ($oShape.LineEndCenter() = $bEndCenter) ? $iError : BitOR($iError, 64)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeGetObjByName
; Description ...: Retrieve a Shape Object, for later Shape related functions.
; Syntax ........: _LOWriter_ShapeGetObjByName(ByRef $oDoc, $sShapeName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sShapeName          - a string value. The Shape name to retrieve the object for.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sShapeName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Draw Page Object
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Shape requested in $sShapeName not found in document.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success, Returning the requested Shape Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ShapesGetNames
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapeGetObjByName(ByRef $oDoc, $sShapeName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShapes

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

		$oShapes = $oDoc.DrawPage()
	If Not IsObj($oShapes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $oShapes.hasElements() Then
		For $i = 0 To $oShapes.getCount() - 1
			If ($oShapes.getByIndex($i).Name() = $sShapeName) Then Return SetError($__LOW_STATUS_SUCCESS, 2, $oShapes.getByIndex($i))

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0);Shape not found
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHasShapeName
; Description ...: Check if a Document contains a Shape with the specified name.
; Syntax ........: _LOWriter_DocHasShapeName(ByRef $oDoc, $sShapeName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sShapeName          - a string value. The Shape name to search for.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sShapeName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Draw Page Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return False = Success. Search was successful, no Shapes found matching $sShapeName.
;				   @Error 0 @Extended 1 Return True = Success. Search was successful, Shape found matching $sShapeName.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ShapeGetObjByName
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_DocHasShapeName(ByRef $oDoc, $sShapeName)
Local $oShapes

	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oShapes = $oDoc.DrawPage()
	If Not IsObj($oShapes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $oShapes.hasElements() Then
		For $i = 0 To $oShapes.getCount() - 1
			If ($oShapes.getByIndex($i).Name() = $sShapeName) Then Return SetError($__LOW_STATUS_SUCCESS,0,True)

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, False) ;No matches
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeInsert
; Description ...: Insert a shape into a document.
; Syntax ........: _LOWriter_ShapeInsert(ByRef $oDoc, ByRef $oCursor, $iShapeType, $iWidth, $iHeight)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions. See Remarks.
;                  $iShapeType          - an integer value (0-122). The Type of shape to create. See remarks. See $LOW_SHAPE_TYPE_* as defined in LibreOfficeWriter_Constants.au3
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iShapeType not an Integer, less than 0, or greater than 122. See $LOW_SHAPE_TYPE_* as defined in LibreOfficeWriter_Constants.au3
;				   @Error 1 @Extended 4 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iHeight not an Integer.
;				   @Error 1 @Extended 6 Return 0 = Cursor called in $oCursor is a Table Cursor, and cannot be used.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve PolyPolygonBezier Structure.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve CustomShapeGeometry Array of Structures.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to determine Cursor type.
;				   @Error 3 @Extended 2 Return 0 = Failed to create requested Shape.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. The Shape was successfully inserted. Returning the Shape's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oCursor cannot be a Table Cursor.
;				   The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;					$LOW_SHAPE_TYPE_ARROWS_ARROW_S_SHAPED, $LOW_SHAPE_TYPE_ARROWS_ARROW_SPLIT, $LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT,
;						$LOW_SHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT
;					$LOW_SHAPE_TYPE_BASIC_CIRCLE_PIE, $LOW_SHAPE_TYPE_BASIC_FRAME
;					$LOW_SHAPE_TYPE_STARS_6_POINT, $LOW_SHAPE_TYPE_STARS_12_POINT, $LOW_SHAPE_TYPE_STARS_SIGNET, $LOW_SHAPE_TYPE_STARS_6_POINT_CONCAVE
;					$LOW_SHAPE_TYPE_SYMBOL_CLOUD, $LOW_SHAPE_TYPE_SYMBOL_FLOWER, $LOW_SHAPE_TYPE_SYMBOL_PUZZLE, $LOW_SHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, $LOW_SHAPE_TYPE_SYMBOL_BEVEL_DIAMOND
;				   Note: Inserting any of the above shapes will still show successful, but the shape will be invisible, and could cause the document to crash.
;				   The following shape is visually different from the manually inserted one in L.O. 7.3.4.2:
;					$LOW_SHAPE_TYPE_SYMBOL_LIGHTNING
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapeInsert(ByRef $oDoc, ByRef $oCursor, $iShapeType, $iWidth, $iHeight)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $iCursorType
	Local $oShape
	Local $tPos, $tPolyCoords
	Local $atCusShapeGeo

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_IntIsBetween($iShapeType,$LOW_SHAPE_TYPE_ARROWS_ARROW_4_WAY,$LOW_SHAPE_TYPE_SYMBOL_PUZZLE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error > 0 Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

Switch $iShapeType

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_4_WAY To $LOW_SHAPE_TYPE_ARROWS_PENTAGON ; Create an Arrow Shape.
		$oShape = __LOWriter_Shape_CreateArrow($oDoc, $iWidth, $iHeight, $iShapeType)
		If @error Then Return SetError($__LW_STATUS_PROCESSING_ERROR,2,0)

$atCusShapeGeo = $oShape.CustomShapeGeometry(); Backup the CustomShapeGeometry property, as it is generally lost upon insertion.
If Not IsArray($atCusShapeGeo) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

	Case $LOW_SHAPE_TYPE_BASIC_ARC To $LOW_SHAPE_TYPE_BASIC_TRIANGLE_RIGHT ; Create a Basic Shape.
		$oShape = __LOWriter_Shape_CreateBasic($oDoc, $iWidth, $iHeight, $iShapeType)
		If @error Then Return SetError($__LW_STATUS_PROCESSING_ERROR,2,0)

If ($iShapeType <> $LOW_SHAPE_TYPE_BASIC_CIRCLE_SEGMENT) And ($iShapeType <> $LOW_SHAPE_TYPE_BASIC_ARC) Then; Arc and Circle Segment shapes are different from the rest, and don't have CustomShapeGeometry property.
	$atCusShapeGeo = $oShape.CustomShapeGeometry(); Backup the CustomShapeGeometry property, as it is generally lost upon insertion.
	If Not IsArray($atCusShapeGeo) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)
EndIf

	Case $LOW_SHAPE_TYPE_CALLOUT_CLOUD To $LOW_SHAPE_TYPE_CALLOUT_ROUND ; Create a Callout Shape.
		$oShape = __LOWriter_Shape_CreateCallout($oDoc, $iWidth, $iHeight, $iShapeType)
		If @error Then Return SetError($__LW_STATUS_PROCESSING_ERROR,2,0)

$atCusShapeGeo = $oShape.CustomShapeGeometry(); Backup the CustomShapeGeometry property, as it is generally lost upon insertion.
If Not IsArray($atCusShapeGeo) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

	Case $LOW_SHAPE_TYPE_FLOWCHART_CARD To $LOW_SHAPE_TYPE_FLOWCHART_TERMINATOR ; Create a Flowchart Shape.
		$oShape = __LOWriter_Shape_CreateFlowchart($oDoc, $iWidth, $iHeight, $iShapeType)
		If @error Then Return SetError($__LW_STATUS_PROCESSING_ERROR,2,0)

$atCusShapeGeo = $oShape.CustomShapeGeometry(); Backup the CustomShapeGeometry property, as it is generally lost upon insertion.
If Not IsArray($atCusShapeGeo) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

	Case $LOW_SHAPE_TYPE_LINE_CURVE To $LOW_SHAPE_TYPE_LINE_POLYGON_45_FILLED ; Create a Line Shape.
		$oShape = __LOWriter_Shape_CreateLine($oDoc, $iWidth, $iHeight, $iShapeType)
		If @error Then Return SetError($__LW_STATUS_PROCESSING_ERROR,2,0)

$tPolyCoords = $oShape.PolyPolygonBezier(); Backup the PolyPolygonBezier property, as it is generally lost upon insertion.
If Not IsObj($tPolyCoords) Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

	Case $LOW_SHAPE_TYPE_STARS_4_POINT To $LOW_SHAPE_TYPE_STARS_SIGNET ; Create a Star or Banner Shape.
		$oShape = __LOWriter_Shape_CreateStars($oDoc, $iWidth, $iHeight, $iShapeType)
		If @error Then Return SetError($__LW_STATUS_PROCESSING_ERROR,2,0)

$atCusShapeGeo = $oShape.CustomShapeGeometry(); Backup the CustomShapeGeometry property, as it is generally lost upon insertion.
If Not IsArray($atCusShapeGeo) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

	Case $LOW_SHAPE_TYPE_SYMBOL_BEVEL_DIAMOND To $LOW_SHAPE_TYPE_SYMBOL_PUZZLE ; Create a Symbol Shape.
		$oShape = __LOWriter_Shape_CreateSymbol($oDoc, $iWidth, $iHeight, $iShapeType)
		If @error Then Return SetError($__LW_STATUS_PROCESSING_ERROR,2,0)

$atCusShapeGeo = $oShape.CustomShapeGeometry(); Backup the CustomShapeGeometry property, as it is generally lost upon insertion.
If Not IsArray($atCusShapeGeo) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

EndSwitch

$tPos = $oShape.Position(); Backup the position, as it is generally lost upon insertion.
If Not IsObj($tPos) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$oCursor.Text.insertTextContent($oCursor,$oShape,False)

$oShape.AnchorType = $LW_ANCHOR_AT_PARAGRAPH

If IsObj($tPolyCoords) Then
	$oShape.PolyPolygonBezier = $tPolyCoords; If shape used the PolyPolyGonBezier property, re-Set it after insertion.

ElseIf IsArray($atCusShapeGeo) Then
$oShape.CustomShapeGeometry = $atCusShapeGeo; If shape used the CustomSHapeGeometry property, re-Set it after insertion.

EndIf

$oShape.Position = $tPos; re-Set the position after insertion.

Return SetError($__LOW_STATUS_SUCCESS,0,$oShape)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeRotateSlant
; Description ...: Set or retrieve Rotation and Slant settings for a Shape.
; Syntax ........: _LOWriter_ShapeRotateSlant(ByRef $oShape[, $nRotate = Null[, $nSlant = Null]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $nRotate             - [optional] a general number value (0-359.99). Default is Null. The Degrees to rotate the shape. See remarks.
;                  $nSlant              - [optional] a general number value (-89-89.00). Default is Null. The Degrees to slant the shape. See remarks.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $nRotate not a Number, less than 0, or greater than 359.99.
;				   @Error 1 @Extended 3 Return 0 = $nSlant not a Number, less than -89, or greater than 89.00.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $nRotate
;				   |								2 = Error setting $nSlant
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function uses the deprecated Libre Office methods RotateAngle, and ShearAngle, and may stop working in future Libre Office versions, after 7.3.4.2.
;				   At the present time Control Point settings are not included as they are too complex to manipulate.
;				   At the present time Corner Radius setting is not included, as I was unable to identify a shape that utilized this setting.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapeRotateSlant(ByRef $oShape, $nRotate = Null, $nSlant = Null)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $aiShape[2]
Local $iError = 0

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

If __LOWriter_VarsAreNull($nRotate, $nSlant) Then
		__LOWriter_ArrayFill($aiShape, ($oShape.RotateAngle()) / 100, ($oShape.ShearAngle()) / 100); Divide by 100 to match L.O. values.
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiShape)
EndIf

If ($nRotate <> Null) Then
		If Not __LOWriter_NumIsBetween($nRotate,0,359.99) Then Return SetError($__LOW_STATUS_INPUT_ERROR,2, 0)
$oShape.RotateAngle = ($nRotate * 100); * 100 to match L.O. Values.
		$iError = (($oShape.RotateAngle() / 100) = $nRotate) ? $iError : BitOR($iError, 1)
EndIf

	If ($nSlant <> Null) Then
		If Not __LOWriter_NumIsBetween($nSlant, -89, 89) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oShape.ShearAngle = ($nSlant * 100); * 100 to match L.O. Values.
		$iError = (($oShape.ShearAngle() / 100) = $nSlant) ? $iError : BitOR($iError, 2)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeDelete
; Description ...: Delete a Shape.
; Syntax ........: _LOWriter_ShapeDelete(ByRef $oDoc, $oShape)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oShape              - an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oShape not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Shape's name.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Shape with the same name still exists in document after deletion attempt.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Shape was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapeDelete(ByRef $oDoc, $oShape)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sShapeName

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

$sShapeName = $oShape.Name()
If Not IsString($sShapeName) Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

$oDoc.getDrawPage().remove($oShape)

	Return (_LOWriter_DocHasShapeName($oDoc,$sShapeName)) ? SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeTextBox
; Description ...: Activate, Set, and Retrieve Shape TextBox settings.
; Syntax ........: _LOWriter_ShapeTextBox(ByRef $oShape[, $bTextBox = Null[, $sContent = Null]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $bTextBox            - [optional] a boolean value. Default is Null. If True, adds a TexttBox inside of the Shape. See Remarks.
;                  $sContent            - [optional] a string value. Default is Null. The Text content of the Shape's TextBox.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bTextBox not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sContent not a String.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Shape called in $oShape does not support "com.sun.star.drawing.CustomShape", and does not support adding a TextBox.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $bTextBox
;				   |								2 = Error setting $sContent
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes do not support adding a TextBox:
;					$LOW_SHAPE_TYPE_LINE_LINE, $LOW_SHAPE_TYPE_LINE_FREEFORM_LINE, $LOW_SHAPE_TYPE_LINE_FREEFORM_LINE_FILLED, $LOW_SHAPE_TYPE_LINE_CURVE, $LOW_SHAPE_TYPE_LINE_CURVE_FILLED,
;						$LOW_SHAPE_TYPE_LINE_POLYGON, $LOW_SHAPE_TYPE_LINE_POLYGON_45, $LOW_SHAPE_TYPE_LINE_POLYGON_45_FILLED.
;					$LOW_SHAPE_TYPE_BASIC_CIRCLE_SEGMENT, $LOW_SHAPE_TYPE_BASIC_ARC.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapeTextBox(ByRef $oShape, $bTextBox = Null, $sContent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avTextBox[2]

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If ($oShape.ShapeType <> "com.sun.star.drawing.CustomShape") Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

If __LOWriter_VarsAreNull($bTextBox, $sContent) Then
		__LOWriter_ArrayFill($avTextBox, $oShape.TextBox(), $oShape.String())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avTextBox)
EndIf

If ($bTextBox <> Null) Then
		If Not IsBool($bTextBox) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oShape.TextBox = $bTextBox
		$iError = ($oShape.TextBox () = $bTextBox ) ? $iError : BitOR($iError, 1)
EndIf

If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oShape.String = $sContent
		$iError = ($oShape.String() = $sContent) ? $iError : BitOR($iError, 2)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeName
; Description ...: Set or Retrieve a Shape's Name.
; Syntax ........: _LOWriter_ShapeName(Byref $oDoc, Byref $oShape[, $sName = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $sName               - [optional] a string value. Default is Null. The new Name for the Shape.
; Return values .: Success: 1 or String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sName not a String.
;				   @Error 1 @Extended 4 Return 0 = Document already contains a Shape with the same name as called in $sName.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sName
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Shape's name was successfully set.
;				   @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning the Shape's current name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapeName(ByRef $oDoc, ByRef $oShape, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($sName = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oShape.Name())

		If Not IsString($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		If _LOWriter_DocHasShapeName($oDoc, $sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oShape.Name = $sName
		$iError = ($oShape.Name() = $sName) ? $iError : BitOR($iError, 1)

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapePosition
; Description ...: Set or Retrieve the Shape's position settings.
; Syntax ........: _LOWriter_ShapePosition(ByRef $oShape[, $iX = Null[, $iY = Null[, $bProtectPos = Null]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $iX                  - [optional] an integer value. Default is Null. The X position from the insertion point, in Micrometers.
;                  $iY                  - [optional] an integer value. Default is Null. The Y position from the insertion point, in Micrometers.
;                  $bProtectPos         - [optional] a boolean value. Default is Null. If True, the Shape's position is locked.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iX not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iY not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $bProtectPos not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Shape's Position Structure.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $iX
;				   |								2 = Error setting $iY
;				   |								4 = Error setting $bProtectPos
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName, _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOWriter_ShapePosition(ByRef $oShape, $iX = Null, $iY = Null, $bProtectPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPosition [3]
	Local $tPos

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1,0)

	If __LOWriter_VarsAreNull($iX, $iY, $bProtectPos) Then
			__LOWriter_ArrayFill($avPosition, $tPos.X(), $tPos.Y(), $oShape.MoveProtect())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avPosition )
	EndIf

If ($iX <> Null) Or ($iY <> Null) Then

	If ($iX <> Null) Then
		If Not IsInt($iX) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$tPos.X = $iX
	EndIf

	If ($iY <> Null) Then
		If Not IsInt($iY) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$tPos.Y = $iY
	EndIf

	$oShape.Position = $tPos

		$iError = ($iX = Null) ? $iError : (__LOWriter_IntIsBetween($oShape.Position.X(), $iX - 1, $iX + 1)) ? $iError : BitOR($iError, 1)
		$iError = ($iY = Null) ? $iError : (__LOWriter_IntIsBetween($oShape.Position.Y(), $iY -1, $iY + 1)) ? $iError : BitOR($iError, 2)
EndIf

	If ($bProtectPos <> Null) Then
		If Not IsBool($bProtectPos ) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oShape.MoveProtect = $bProtectPos
		$iError = ($oShape.MoveProtect() = $bProtectPos) ? $iError : BitOR($iError, 4)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeTransparency
; Description ...: Set or retrieve Transparency settings for a Shape.
; Syntax ........: _LOWriter_ShapeTransparency(Byref $oShape[, $iTransparency = Null])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. The color transparency. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iTransparency
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;				   @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current setting for Transparency in integer format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapeTransparency(ByRef $oShape, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iTransparency) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oShape.FillTransparence())

	If Not __LOWriter_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oShape.FillTransparenceGradientName = "" ;Turn of Gradient if it is on, else settings wont be applied.
	$oShape.FillTransparence = $iTransparency
	$iError = ($oShape.FillTransparence() = $iTransparency) ? $iError : BitOR($iError, 1)

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeTransparencyGradient
; Description ...: Set or retrieve the Shape transparency gradient settings.
; Syntax ........: _LOWriter_ShapeTransparencyGradient(Byref $oDoc, Byref $oShape[, $iType = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iBorder = Null[, $iStart = Null[, $iEnd = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The type of transparency gradient that you want to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_GRAD_TYPE_OFF to turn Transparency Gradient off.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iBorder             - [optional] an integer value (0-100). Default is Null. The amount by which you want to adjust the transparent area of the gradient. Set in percentage.
;                  $iStart              - [optional] an integer value (0-100). Default is Null. The transparency value for the beginning point of the gradient, where 0% is fully opaque and 100% is fully transparent.
;                  $iEnd                - [optional] an integer value (0-100). Default is Null. The transparency value for the endpoint of the gradient, where 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iType not an Integer, less than -1, or greater than 5, see constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $iXCenter not an Integer, less than 0, or greater than 100.
;				   @Error 1 @Extended 5 Return 0 = $iYCenter not an Integer, less than 0, or greater than 100.
;				   @Error 1 @Extended 6 Return 0 = $iAngle not an Integer, less than 0, or greater than 359.
;				   @Error 1 @Extended 7 Return 0 = $iBorder not an Integer, less than 0, or greater than 100.
;				   @Error 1 @Extended 8 Return 0 = $iStart not an Integer, less than 0, or greater than 100.
;				   @Error 1 @Extended 9 Return 0 = $iEnd not an Integer, less than 0, or greater than 100.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving "FillTransparenceGradient" Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.drawing.TransparencyGradientTable" Object.
;				   @Error 2 @Extended 3 Return 0 = Error creating "com.sun.star.awt.Gradient" structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error creating Transparency Gradient Name.
;				   @Error 3 @Extended 2 Return 0 = Error setting Transparency Gradient Name.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iType
;				   |								2 = Error setting $iXCenter
;				   |								4 = Error setting $iYCenter
;				   |								8 = Error setting $iAngle
;				   |								16 = Error setting $iBorder
;				   |								32 = Error setting $iStart
;				   |								64 = Error setting $iEnd
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;				   @Error 0 @Extended 0 Return 2 = Success. Transparency Gradient has been successfully turned off.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapeTransparencyGradient(ByRef $oDoc, ByRef $oShape, $iType = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iBorder = Null, $iStart = Null, $iEnd = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tGradient
	Local $sTGradName
	Local $iError = 0
	Local $aiTransparent[7]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$tGradient = $oShape.FillTransparenceGradient()
	If Not IsObj($tGradient) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iType, $iXCenter, $iYCenter, $iAngle, $iBorder, $iStart, $iEnd) Then
		__LOWriter_ArrayFill($aiTransparent, $tGradient.Style(), $tGradient.XOffset(), $tGradient.YOffset(), _
				($tGradient.Angle() / 10), $tGradient.Border(), __LOWriter_TransparencyGradientConvert(Null, $tGradient.StartColor()), _
				__LOWriter_TransparencyGradientConvert(Null, $tGradient.EndColor())) ; Angle is set in thousands
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiTransparent)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oShape.FillTransparenceGradientName = ""
			Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LOWriter_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$tGradient.Style = $iType
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LOWriter_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$tGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LOWriter_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$tGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LOWriter_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$tGradient.Angle = ($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iBorder <> Null) Then
		If Not __LOWriter_IntIsBetween($iBorder, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tGradient.Border = $iBorder
	EndIf

	If ($iStart <> Null) Then
		If Not __LOWriter_IntIsBetween($iStart, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$tGradient.StartColor = __LOWriter_TransparencyGradientConvert($iStart)
	EndIf

	If ($iEnd <> Null) Then
		If Not __LOWriter_IntIsBetween($iEnd, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$tGradient.EndColor = __LOWriter_TransparencyGradientConvert($iEnd)
	EndIf

	If ($oShape.FillTransparenceGradientName() = "") Then
		$sTGradName = __LOWriter_TransparencyGradientNameInsert($oDoc, $tGradient)
		If @error > 0 Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

		$oShape.FillTransparenceGradientName = $sTGradName
		If ($oShape.FillTransparenceGradientName <> $sTGradName) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
	EndIf

	$oShape.FillTransparenceGradient = $tGradient

	$iError = ($iType = Null) ? $iError : ($oShape.FillTransparenceGradient.Style() = $iType) ? $iError : BitOR($iError, 1)
	$iError = ($iXCenter = Null) ? $iError : ($oShape.FillTransparenceGradient.XOffset() = $iXCenter) ? $iError : BitOR($iError, 2)
	$iError = ($iYCenter = Null) ? $iError : ($oShape.FillTransparenceGradient.YOffset() = $iYCenter) ? $iError : BitOR($iError, 4)
	$iError = ($iAngle = Null) ? $iError : (($oShape.FillTransparenceGradient.Angle() / 10) = $iAngle) ? $iError : BitOR($iError, 8)
	$iError = ($iBorder = Null) ? $iError : ($oShape.FillTransparenceGradient.Border() = $iBorder) ? $iError : BitOR($iError, 16)
	$iError = ($iStart = Null) ? $iError : ($oShape.FillTransparenceGradient.StartColor() = __LOWriter_TransparencyGradientConvert($iStart)) ? $iError : BitOR($iError, 32)
	$iError = ($iEnd = Null) ? $iError : ($oShape.FillTransparenceGradient.EndColor() = __LOWriter_TransparencyGradientConvert($iEnd)) ? $iError : BitOR($iError, 64)

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeTypePosition
; Description ...: Set or Retrieve Shape Position Settings.
; Syntax ........: _LOWriter_ShapeTypePosition(Byref $oShape[, $iHorAlign = Null[, $iHorPos = Null[, $iHorRelation = Null[, $bMirror = Null[, $iVertAlign = Null[, $iVertPos = Null[, $iVertRelation = Null[, $bKeepInside = Null[, $iAnchorPos = Null]]]]]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $iHorAlign           - [optional] an integer value (0-3). Default is Null. The horizontal orientation of the Shape. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3. Can't be set if Anchor position is set to "As Character".
;                  $iHorPos             - [optional] an integer value. Default is Null. The horizontal position of the Shape. set in Micrometer(uM). Only valid if $iHorAlign is set to $LOW_ORIENT_HORI_NONE().
;                  $iHorRelation        - [optional] an integer value (0-8). Default is Null. The reference point for the selected horizontal alignment option. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3, and Remarks for acceptable values.
;                  $bMirror             - [optional] a boolean value. Default is Null. If True, Reverses the current horizontal alignment settings on even pages.
;                  $iVertAlign          - [optional] an integer value (0-9). Default is Null. The vertical orientation of the Shape. See Constants, $LOW_ORIENT_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertPos            - [optional] an integer value. Default is Null. The vertical position of the Shape. set in Micrometer(uM). Only valid if $iVertAlign is set to $LOW_ORIENT_VERT_NONE().
;                  $iVertRelation       - [optional] an integer value (-1-9). Default is Null. The reference point for the selected vertical alignment option. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3, and Remarks for acceptable values.
;                  $bKeepInside         - [optional] a boolean value. Default is Null. If True, Keeps the Shape within the layout boundaries of the text that the Shape is anchored to.
;                  $iAnchorPos          - [optional] an integer value(0,1,4). Default is Null. Specify the anchoring options for the Shape. See Constants, $LOW_ANCHOR_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iHorAlign not an Integer, or less than 0, or greater than 3. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iHorPos not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iHorRelation not an Integer, or less than 0, or greater than 8. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 5 Return 0 = $bMirror not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iVertAlign not an integer, or less than 0, or greater than 9. See Constants, $LOW_ORIENT_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 7 Return 0 = $iVertPos not an integer.
;				   @Error 1 @Extended 8 Return 0 = $iVertRelation not an Integer, Less than -1, or greater than 9. See Constants, $LOW_RELATIVE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 9 Return 0 = $bKeepInside not a Boolean.
;				   @Error 1 @Extended 10 Return 0 = $iAnchorPos not an Integer, or not equal to 0, 1 or 4. See Constants, $LOW_ANCHOR_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iHorAlign
;				   |								2 = Error setting $iHorPos
;				   |								4 = Error setting $iHorRelation
;				   |								8 = Error setting $bMirror
;				   |								16 = Error setting $iVertAlign
;				   |								32 = Error setting $iVertPos
;				   |								64 = Error setting $iVertRelation
;				   |								128 = Error setting $bKeepInside
;				   |								256 = Error setting $iAnchorPos
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 9 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   $iHorRelation has varying acceptable values, depending on the current Anchor position and also the current
;							$iHorAlign setting. The Following is a list of acceptable values per anchor position.
;						$LOW_ANCHOR_AT_PARAGRAPH(0) Accepts the following $iHorRelation Values:
;							$LOW_RELATIVE_PARAGRAPH (0),
;							$LOW_RELATIVE_PARAGRAPH_TEXT (1),
;							$LOW_RELATIVE_PAGE_LEFT (3)[Same as Left Page Border in L.O. UI],
;							$LOW_RELATIVE_PAGE_RIGHT (4)[Same as Right Page Border in L.O. UI],
;							$LOW_RELATIVE_PARAGRAPH_LEFT (5),
;							$LOW_RELATIVE_PARAGRAPH_RIGHT (6),
;							$LOW_RELATIVE_PAGE (7),
;							$LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;						$LOW_ANCHOR_AS_CHARACTER(1) Accepts No $iHorRelation Values.
;						$LOW_ANCHOR_AT_CHARACTER(4) Accepts the following $iHorRelation Values:
;							$LOW_RELATIVE_PARAGRAPH (0),
;							$LOW_RELATIVE_PARAGRAPH_TEXT (1),
;							$LOW_RELATIVE_CHARACTER (2),
;							$LOW_RELATIVE_PAGE_LEFT (3)[Same as Left Page Border in L.O. UI],
;							$LOW_RELATIVE_PAGE_RIGHT (4)[Same as Right Page Border in L.O. UI],
;							$LOW_RELATIVE_PARAGRAPH_LEFT (5),
;							$LOW_RELATIVE_PARAGRAPH_RIGHT (6),
;							$LOW_RELATIVE_PAGE (7),
;							$LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;					$iVertRelation has varying acceptable values, depending on the current Anchor position. The Following is a list of acceptable values per anchor position.
;						$LOW_ANCHOR_AT_PARAGRAPH(0) Accepts the following $iVertRelation Values:
;							$LOW_RELATIVE_PARAGRAPH (0)[The Same as "Margin" in L.O. UI],
;							$LOW_RELATIVE_PAGE (7),
;							$LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;						$LOW_ANCHOR_AS_CHARACTER(1) Accepts the following $iVertRelation Values:
;							$LOW_RELATIVE_ROW(-1),
;							$LOW_RELATIVE_PARAGRAPH (0)[The Same as "Baseline" in L.O. UI],
;							$LOW_RELATIVE_CHARACTER (2),
;						$LOW_ANCHOR_AT_CHARACTER(4) Accepts the following $iVertRelation Values:
;							$LOW_RELATIVE_PARAGRAPH (0)[The same as "Margin" in L.O. UI],
;							$LOW_RELATIVE_PARAGRAPH_TEXT (1),
;							$LOW_RELATIVE_CHARACTER (2),
;							$LOW_RELATIVE_PAGE (7),
;							$LOW_RELATIVE_PAGE_PRINT (8)[Same as Page Text Area in L.O. UI].
;							$LOW_RELATIVE_TEXT_LINE (9)[The same as "Line of Text" in L.O. UI]
;					The behaviour of each Relation constant is described below.
;							$LOW_RELATIVE_ROW(-1), This option will position the Shape considering the height of the row where the anchor is placed.
;							$LOW_RELATIVE_PARAGRAPH (0), [For Horizontal Relation:] the Shape is positioned considering the whole width available for the paragraph, including indent spaces.
;								[$LOW_RELATIVE_PARAGRAPH for Vertical Relation:] {$LOW_RELATIVE_PARAGRAPH is Also called "Margin" or "Baseline" in L.O. UI], Depending on the anchoring type, the Shape is positioned considering the space between the top margin and the character ("To character" anchoring) or bottom edge of the paragraph ("To paragraph" anchoring) where the anchor is placed. Or will position the Shape considering the text baseline over which all characters are placed. ("As Character" anchoring.)
;							$LOW_RELATIVE_PARAGRAPH_TEXT (1), [For Horizontal Relation:] the Shape is positioned considering the whole width available for text in the paragraph, excluding indent spaces.
;								[$LOW_RELATIVE_PARAGRAPH_TEXT for Vertical relation:] the Shape is positioned considering the height of the paragraph where the anchor is placed.
;							$LOW_RELATIVE_CHARACTER (2), [For Horizontal Relation:] the Shape is positioned considering the horizontal space used by the character.
;								[$LOW_RELATIVE_CHARACTER for Vertical relation:] the Shape is positioned considering the vertical space used by the character.
;							$LOW_RELATIVE_PAGE_LEFT (3),[For Horizontal Relation:], the Shape is positioned considering the space available between the left page border and the left paragraph border. [Same as Left Page Border in L.O. UI]
;							$LOW_RELATIVE_PAGE_RIGHT (4),[For Horizontal Relation:], the Shape is positioned considering the space available between the Right page border and the right paragraph border. [Same as Right Page Border in L.O. UI]
;							$LOW_RELATIVE_PARAGRAPH_LEFT (5),[For Horizontal Relation:] the Shape is positioned considering the width of the indent space available to the left of the paragraph.
;							$LOW_RELATIVE_PARAGRAPH_RIGHT (6),[For Horizontal Relation:], the Shape is positioned considering the width of the indent space available to the right of the paragraph.
;							$LOW_RELATIVE_PAGE (7),[For Horizontal Relation:], the Shape is positioned considering the whole width of the page, from the left to the right page borders.
;								[$LOW_RELATIVE_PAGE for Vertical relation:], the Shape is positioned considering the full page height, from top to bottom page borders.
;							$LOW_RELATIVE_PAGE_PRINT (8),[For Horizontal Relation:], [Same as Page Text Area in L.O. UI] the Shape is positioned considering the whole width available for text in the page, from the left to the right page margins.
;								[$LOW_RELATIVE_PAGE_PRINT for Vertical relation:], the Shape is positioned considering the full height available for text, from top to bottom margins.
;							$LOW_RELATIVE_TEXT_LINE (9),[For Vertical relation:], the Shape is positioned considering the height of the line of text where the anchor is placed.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapeTypePosition(ByRef $oShape, $iHorAlign = Null, $iHorPos = Null, $iHorRelation = Null, $bMirror = Null, $iVertAlign = Null, $iVertPos = Null, $iVertRelation = Null, $bKeepInside = Null, $iAnchorPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurrentAnchor
	Local $avPosition[9]

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iHorAlign, $iHorPos, $iHorRelation, $bMirror, $iVertAlign, $iVertPos, $iVertRelation, $bKeepInside, $iAnchorPos) Then
		__LOWriter_ArrayFill($avPosition, $oShape.HoriOrient(), $oShape.HoriOrientPosition(), $oShape.HoriOrientRelation(), _
				$oShape.PageToggle(), $oShape.VertOrient(), $oShape.VertOrientPosition(), $oShape.VertOrientRelation(), _
				$oShape.IsFollowingTextFlow(), $oShape.AnchorType())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avPosition)
	EndIf
	; Accepts HoriOrient Left,Right, Center, and "None" = "From Left"
	If ($iHorAlign <> Null) Then ; Cant be set if Anchor is set to "As Char"
		If Not __LOWriter_IntIsBetween($iHorAlign, $LOW_ORIENT_HORI_NONE, $LOW_ORIENT_HORI_LEFT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oShape.HoriOrient = $iHorAlign
		$iError = ($oShape.HoriOrient() = $iHorAlign) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iHorPos <> Null) Then
		If Not IsInt($iHorPos) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oShape.HoriOrientPosition = $iHorPos
		$iError = (__LOWriter_IntIsBetween($oShape.HoriOrientPosition(), $iHorPos - 1, $iHorPos + 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iHorRelation <> Null) Then
		If Not __LOWriter_IntIsBetween($iHorRelation, $LOW_RELATIVE_PARAGRAPH, $LOW_RELATIVE_PAGE_PRINT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oShape.HoriOrientRelation = $iHorRelation
		$iError = ($oShape.HoriOrientRelation() = $iHorRelation) ? $iError : BitOR($iError, 4)
	EndIf

	If ($bMirror <> Null) Then
		If Not IsBool($bMirror) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oShape.PageToggle = $bMirror
		$iError = ($oShape.PageToggle() = $bMirror) ? $iError : BitOR($iError, 8)
	EndIf

	; Accepts Orient Top,Bottom, Center, and "None" = "From Top"/From Bottom, plus Row and Char.
	If ($iVertAlign <> Null) Then
		If Not __LOWriter_IntIsBetween($iVertAlign, $LOW_ORIENT_VERT_NONE, $LOW_ORIENT_VERT_LINE_BOTTOM) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oShape.VertOrient = $iVertAlign
		$iError = ($oShape.VertOrient() = $iVertAlign) ? $iError : BitOR($iError, 16)
	EndIf

	If ($iVertPos <> Null) Then
		If Not IsInt($iVertPos) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oShape.VertOrientPosition = $iVertPos
		$iError = (__LOWriter_IntIsBetween($oShape.VertOrientPosition(), $iVertPos - 1, $iVertPos + 1)) ? $iError : BitOR($iError, 32)
	EndIf

	If ($iVertRelation <> Null) Then
		If Not __LOWriter_IntIsBetween($iVertRelation, $LOW_RELATIVE_ROW, $LOW_RELATIVE_TEXT_LINE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$iCurrentAnchor = (($iAnchorPos <> Null) ? $iAnchorPos : $oShape.AnchorType())

		; Libre Office is a bit complex in this anchor setting; When set to "As Character", there aren't specific setting
		;		values for "Baseline, "Character" and "Row", But For Baseline the VertOrientRelation value is 0, or
		; "$LOW_RELATIVE_PARAGRAPH", For "Character", The VertOrientRelation value is still 0, and the "VertOrient" value (In the
		; L.O. UI the furthest left drop down box)  is modified, which can be either $LOW_ORIENT_VERT_CHAR_TOP(1),
		; $LOW_ORIENT_VERT_CHAR_CENTER(2), $LOW_ORIENT_VERT_CHAR_BOTTOM(3), depending on the current value of Top, Bottom and
		; Center, or "From Bottom"/ "From Top", of "VertOrient". The same is true For "Row", which means when the anchor is set
		; to "As Character", I need to first determine the the desired user setting, $LOW_RELATIVE_ROW(-1),
		; $LOW_RELATIVE_PARAGRAPH(0), or $LOW_RELATIVE_CHARACTER(2), and then determine the current "VertOrient" setting, and
		; then manually set the value to the correct setting. Such as Line_Top, Line_Bottom etc.

		If ($iCurrentAnchor = $LOW_ANCHOR_AS_CHARACTER) Then

			If ($iVertRelation = $LOW_RELATIVE_ROW) Then
				Switch $oShape.VertOrient()
					Case $LOW_ORIENT_VERT_NONE ; None = "From Bottom or From Top in L.O. UI
						$iError = BitOR($iError, 64) ; -- Row not accepted with this VertOrient Setting.
					Case $LOW_ORIENT_VERT_TOP, $LOW_ORIENT_VERT_CHAR_TOP, $LOW_ORIENT_VERT_LINE_TOP
						$oShape.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oShape.VertOrient = $LOW_ORIENT_VERT_LINE_TOP
						$iError = (($oShape.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oShape.VertOrient() = $LOW_ORIENT_VERT_LINE_TOP)) ? $iError : BitOR($iError, 64)
					Case $LOW_ORIENT_VERT_CENTER, $LOW_ORIENT_VERT_CHAR_CENTER, $LOW_ORIENT_VERT_LINE_CENTER
						$oShape.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oShape.VertOrient = $LOW_ORIENT_VERT_LINE_CENTER
						$iError = (($oShape.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oShape.VertOrient() = $LOW_ORIENT_VERT_LINE_CENTER)) ? $iError : BitOR($iError, 64)
					Case $LOW_ORIENT_VERT_BOTTOM, $LOW_ORIENT_VERT_CHAR_BOTTOM, $LOW_ORIENT_VERT_LINE_BOTTOM
						$oShape.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oShape.VertOrient = $LOW_ORIENT_VERT_LINE_BOTTOM
						$iError = (($oShape.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oShape.VertOrient() = $LOW_ORIENT_VERT_LINE_BOTTOM)) ? $iError : BitOR($iError, 64)
				EndSwitch

			ElseIf ($iVertRelation = $LOW_RELATIVE_PARAGRAPH) Then ; Paragraph = Baseline setting in L.O. UI
				$oShape.VertOrientRelation = $iVertRelation ;Paragraph = Baseline in this case
				$iError = (($oShape.VertOrientRelation() = $iVertRelation)) ? $iError : BitOR($iError, 64)
			ElseIf ($iVertRelation = $LOW_RELATIVE_CHARACTER) Then
				Switch $oShape.VertOrient()
					Case $LOW_ORIENT_VERT_NONE ; None = "From Bottom or From Top in L.O. UI
						$iError = BitOR($iError, 64) ; -- Character not accepted with this VertOrient Setting.
					Case $LOW_ORIENT_VERT_TOP, $LOW_ORIENT_VERT_CHAR_TOP, $LOW_ORIENT_VERT_LINE_TOP
						$oShape.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oShape.VertOrient = $LOW_ORIENT_VERT_CHAR_TOP
						$iError = (($oShape.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oShape.VertOrient() = $LOW_ORIENT_VERT_CHAR_TOP)) ? $iError : BitOR($iError, 64)
					Case $LOW_ORIENT_VERT_CENTER, $LOW_ORIENT_VERT_CHAR_CENTER, $LOW_ORIENT_VERT_LINE_CENTER
						$oShape.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oShape.VertOrient = $LOW_ORIENT_VERT_CHAR_CENTER
						$iError = (($oShape.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oShape.VertOrient() = $LOW_ORIENT_VERT_CHAR_CENTER)) ? $iError : BitOR($iError, 64)
					Case $LOW_ORIENT_VERT_BOTTOM, $LOW_ORIENT_VERT_CHAR_BOTTOM, $LOW_ORIENT_VERT_LINE_BOTTOM
						$oShape.VertOrientRelation = $LOW_RELATIVE_PARAGRAPH
						$oShape.VertOrient = $LOW_ORIENT_VERT_CHAR_BOTTOM
						$iError = (($oShape.VertOrientRelation() = $LOW_RELATIVE_PARAGRAPH) And ($oShape.VertOrient() = $LOW_ORIENT_VERT_CHAR_BOTTOM)) ? $iError : BitOR($iError, 64)
				EndSwitch
			EndIf

		Else
			$oShape.VertOrientRelation = $iVertRelation
			$iError = ($oShape.VertOrientRelation() = $iVertRelation) ? $iError : BitOR($iError, 64)
		EndIf
	EndIf

	If ($bKeepInside <> Null) Then
		If Not IsBool($bKeepInside) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$oShape.IsFollowingTextFlow = $bKeepInside
		$iError = ($oShape.IsFollowingTextFlow() = $bKeepInside) ? $iError : BitOR($iError, 128)
	EndIf

	If ($iAnchorPos <> Null) Then
		If Not __LOWriter_IntIsBetween($iAnchorPos, $LOW_ANCHOR_AT_PARAGRAPH, $LOW_ANCHOR_AS_CHARACTER, "", $LOW_ANCHOR_AT_CHARACTER ) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
		$oShape.AnchorType = $iAnchorPos
		$iError = ($oShape.AnchorType() = $iAnchorPos) ? $iError : BitOR($iError, 256)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeTypeSize
; Description ...: Set or Retrieve Shape Size related settings.
; Syntax ........: _LOWriter_ShapeTypeSize(ByRef $oShape[, $iWidth = Null[, $iHeight = Null[, $bProtectSize = Null]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the Shape, in Micrometers(uM). Min. 51.
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the Shape, in Micrometers(uM). Min. 51.
;                  $bProtectSize        - [optional] a boolean value. Default is Null. If True, Locks the size of the Shape.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iWidth not an Integer, or less than 51.
;				   @Error 1 @Extended 3 Return 0 = $iHeight not an Integer, or less than 51.
;				   @Error 1 @Extended 4 Return 0 = $bProtectSize not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iWidth
;				   |								2 = Error setting $iHeight
;				   |								4 = Error setting $bProtectSize
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   I have skipped "Keep Ratio, as currently it seems unable to be set for shapes.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName, _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapeTypeSize(ByRef $oShape, $iWidth = Null, $iHeight = Null, $bProtectSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSize[3]
	Local $tSize

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1,0)

	If __LOWriter_VarsAreNull($iWidth, $iHeight, $bProtectSize) Then
			__LOWriter_ArrayFill($avSize, $tSize.Width(), $tSize.Height(), $oShape.SizeProtect())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avSize)
	EndIf

If ($iWidth <> Null) Or ($iHeight <> Null) Then

	If ($iWidth <> Null) Then ; Min 51
		If Not __LOWriter_IntIsBetween($iWidth, 51, $iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$tSize.Width = $iWidth
	EndIf

	If ($iHeight <> Null) Then
		If Not __LOWriter_IntIsBetween($iHeight, 51, $iHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$tSize.Height = $iHeight
	EndIf

	$oShape.Size = $tSize

		$iError = ($iWidth = Null) ? $iError : (__LOWriter_IntIsBetween($oShape.Size.Width(), $iWidth - 1, $iWidth + 1)) ? $iError : BitOR($iError, 1)
		$iError = ($iHeight = Null) ? $iError : (__LOWriter_IntIsBetween($oShape.Size.Height(), $iHeight -1, $iHeight + 1)) ? $iError : BitOR($iError, 2)
EndIf

	If ($bProtectSize <> Null) Then
		If Not IsBool($bProtectSize) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oShape.SizeProtect = $bProtectSize
		$iError = ($oShape.SizeProtect() = $bProtectSize) ? $iError : BitOR($iError, 4)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeWrap
; Description ...: Set or Retrieve Shape Wrap and Spacing settings.
; Syntax ........: _LOWriter_ShapeWrap(Byref $oShape[, $iWrapType = Null[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $iWrapType           - [optional] an integer value (0-5). Default is Null. The way you want text to wrap around the Shape. See Constants, $LOW_WRAP_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value. Default is Null. The amount of space between the left edge of the Shape and the text. Set in Micrometers.
;                  $iRight              - [optional] an integer value. Default is Null. The amount of space between the Right edge of the Shape and the text. Set in Micrometers.
;                  $iTop                - [optional] an integer value. Default is Null. The amount of space between the Top edge of the Shape and the text. Set in Micrometers.
;                  $iBottom             - [optional] an integer value. Default is Null. The amount of space between the Bottom edge of the Shape and the text. Set in Micrometers.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iWrapType not an Integer, less than 0, or greater than 5. See Constants, $LOW_WRAP_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iLeft not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iRight not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Property Set Info Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iWrapType
;				   |								2 = Error setting $iLeft
;				   |								4 = Error setting $iRight
;				   |								8 = Error setting $iTop
;				   |								16 = Error setting $iBottom
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName, _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapeWrap(ByRef $oShape, $iWrapType = Null, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPropInfo
	Local $iError = 0
	Local $avWrap[5]

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oPropInfo = $oShape.getPropertySetInfo()
	If Not IsObj($oPropInfo) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iWrapType, $iLeft, $iRight, $iTop, $iBottom) Then

		If $oPropInfo.hasPropertyByName("Surround") Then ; Surround is marked as deprecated, but there is no indication of what version of L.O. this occurred. So Test for its existence.
			__LOWriter_ArrayFill($avWrap, $oShape.Surround(), $oShape.LeftMargin(), $oShape.RightMargin(), $oShape.TopMargin(), _
					$oShape.BottomMargin())
		Else
			__LOWriter_ArrayFill($avWrap, $oShape.TextWrap(), $oShape.LeftMargin(), $oShape.RightMargin(), $oShape.TopMargin(), _
					$oShape.BottomMargin())
		EndIf

		Return SetError($__LOW_STATUS_SUCCESS, 1, $avWrap)
	EndIf

	If ($iWrapType <> Null) Then
		If Not __LOWriter_IntIsBetween($iWrapType, $LOW_WRAP_MODE_NONE, $LOW_WRAP_MODE_RIGHT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		If $oPropInfo.hasPropertyByName("Surround") Then $oShape.Surround = $iWrapType
		If $oPropInfo.hasPropertyByName("TextWrap") Then $oShape.TextWrap = $iWrapType

		If $oPropInfo.hasPropertyByName("Surround") Then
			$iError = ($oShape.Surround() = $iWrapType) ? $iError : BitOR($iError, 1)
		Else
			$iError = ($oShape.TextWrap() = $iWrapType) ? $iError : BitOR($iError, 1)
		EndIf

	EndIf

	If ($iLeft <> Null) Then
		If Not IsInt($iLeft) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oShape.LeftMargin = $iLeft
		$iError = (__LOWriter_IntIsBetween($oShape.LeftMargin(), $iLeft - 1, $iLeft + 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iRight <> Null) Then
		If Not IsInt($iRight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oShape.RightMargin = $iRight
		$iError = (__LOWriter_IntIsBetween($oShape.RightMargin(), $iRight - 1, $iRight + 1)) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iTop <> Null) Then
		If Not IsInt($iTop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oShape.TopMargin = $iTop
		$iError = (__LOWriter_IntIsBetween($oShape.TopMargin(), $iTop - 1, $iTop + 1)) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iBottom <> Null) Then
		If Not IsInt($iBottom) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oShape.BottomMargin = $iBottom
		$iError = (__LOWriter_IntIsBetween($oShape.BottomMargin(), $iBottom - 1, $iBottom + 1)) ? $iError : BitOR($iError, 16)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeWrapOptions
; Description ...: Set or Retrieve Shape Wrap Options.
; Syntax ........: _LOWriter_ShapeWrapOptions(Byref $oShape[, $bFirstPar = Null[, $bInBackground = Null[, $bAllowOverlap = Null]]])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $bFirstPar           - [optional] a boolean value. Default is Null. If True, Starts a new paragraph below the object.
;                  $bInBackground       - [optional] a boolean value. Default is Null. If True, Moves the selected object to the background. This option is only available with the "Through" wrap type.
;                  $bAllowOverlap       - [optional] a boolean value. Default is Null. If True, the object is allowed to overlap another object. This option has no effect on wrap through objects, which can always overlap.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bFirstPar not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bInBackground not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bAllowOverlap not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bFirstPar
;				   |								2 = Error setting $bInBackground
;				   |								4 = Error setting $bAllowOverlap
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   This function may indicate the settings were set successfully when they haven't been if the appropriate wrap type, anchor type etc. hasn't been set before hand.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapeWrapOptions(ByRef $oShape, $bFirstPar = Null, $bInBackground = Null, $bAllowOverlap = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abWrapOptions[3]

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bFirstPar, $bInBackground, $bAllowOverlap) Then
		__LOWriter_ArrayFill($abWrapOptions, $oShape.SurroundAnchorOnly(), (($oShape.Opaque()) ? False : True), $oShape.AllowOverlap())
		; Opaque/Background is False when InBackground is checked, so switch Boolean values around.
		Return SetError($__LOW_STATUS_SUCCESS, 1, $abWrapOptions)
	EndIf

	If ($bFirstPar <> Null) Then
		If Not IsBool($bFirstPar) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oShape.SurroundAnchorOnly = $bFirstPar
		$iError = ($oShape.SurroundAnchorOnly() = $bFirstPar) ? $iError : BitOR($iError, 1)
	EndIf

	If ($bInBackground <> Null) Then
		If Not IsBool($bInBackground) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oShape.Opaque = (($bInBackground) ? False : True)
		$iError = ($oShape.Opaque() = (($bInBackground) ? False : True)) ? $iError : BitOR($iError, 2) ; Opaque/Background is False when InBackground is checked, so switch Boolean values around.
	EndIf

	If ($bAllowOverlap <> Null) Then
		If Not IsBool($bAllowOverlap) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oShape.AllowOverlap = $bAllowOverlap
		$iError = ($oShape.AllowOverlap() = $bAllowOverlap) ? $iError : BitOR($iError, 4)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeAreaColor
; Description ...: Set or Retrieve the Fill color settings for a Shape.
; Syntax ........: _LOWriter_ShapeAreaColor(Byref $oShape[, $iColor = Null])
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is Null. The Fill color. Set in Long integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for "None".
; Return values .: Success: 1 or Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iColor not an integer, less than -1, or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iColor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current Fill color as an integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Note: If transparency is set, it can cause strange values to be displayed for Background color.
; Related .......:  _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName, _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapeAreaColor(ByRef $oShape, $iColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

; If $iColor is Null, and Fill Style is set to solid, then return current color value, else return LOW_COLOR_OFF.
	If ($iColor = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, ($oShape.FillStyle() = $__LOWCONST_FILL_STYLE_SOLID) ? $oShape.FillColor() : $LOW_COLOR_OFF)

	If Not __LOWriter_IntIsBetween($iColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If ($iColor = $LOW_COLOR_OFF) Then
	$oShape.FillStyle = $__LOWCONST_FILL_STYLE_OFF
	Else
		$oShape.FillStyle = $__LOWCONST_FILL_STYLE_SOLID
		$oShape.FillColor = $iColor
		$iError = ($oShape.FillColor() = $iColor) ? $iError : BitOR($iError, 1)
		EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeAreaGradient
; Description ...: Modify or retrieve the settings for Shape BackGround color Gradient.
; Syntax ........: _LOWriter_ShapeAreaGradient(Byref $oDoc, Byref $oShape[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iBorder = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
;                  $sGradientName       - [optional] a string value. Default is Null. A Preset Gradient Name. See remarks. See constants, $LOW_GRAD_NAME_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iType               - [optional] an integer value (-1-5). Default is Null. The gradient type to apply. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iIncrement          - [optional] an integer value (0, 3-256). Default is Null. The number of steps of color change. 0 = Automatic.
;                  $iXCenter            - [optional] an integer value (0-100). Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] an integer value (0-100). Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] an integer value (0-359). Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iBorder             - [optional] an integer value (0-100). Default is Null. The amount by which to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] an integer value (0-16777215). Default is Null. A color for the beginning point of the gradient, set in Long Color Integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iToColor            - [optional] an integer value (0-16777215). Default is Null. A color for the endpoint of the gradient, set in Long Color Integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iFromIntense        - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] an integer value (0-100). Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oShape not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sGradientName not a String.
;				   @Error 1 @Extended 4 Return 0 = $iType not an Integer, less than -1, or greater than 5. See Constants, $LOW_GRAD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 5 Return 0 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;				   @Error 1 @Extended 6 Return 0 = $iXCenter not an Integer, less than 0, or greater than 100.
;				   @Error 1 @Extended 7 Return 0 = $iYCenter not an Integer, less than 0, or greater than 100.
;				   @Error 1 @Extended 8 Return 0 = $iAngle not an Integer, less than 0, or greater than 359.
;				   @Error 1 @Extended 9 Return 0 = $iBorder not an Integer, less than 0, or greater than 100.
;				   @Error 1 @Extended 10 Return 0 = $iFromColor not an Integer, less than 0, or greater than 16777215.
;				   @Error 1 @Extended 11 Return 0 = $iToColor not an Integer, less than 0, or greater than 16777215.
;				   @Error 1 @Extended 12 Return 0 = $iFromIntense not an Integer, less than 0, or greater than 100.
;				   @Error 1 @Extended 13 Return 0 = $iToIntense not an Integer, less than 0, or greater than 100.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving "FillGradient" Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error creating Transparency Gradient Name.
;				   @Error 3 @Extended 2 Return 0 = Error setting Transparency Gradient Name.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sGradientName
;				   |								2 = Error setting $iType
;				   |								4 = Error setting $iIncrement
;				   |								8 = Error setting $iXCenter
;				   |								16 = Error setting $iYCenter
;				   |								32 = Error setting $iAngle
;				   |								64 = Error setting $iBorder
;				   |								128 = Error setting $iFromColor
;				   |								256 = Error setting $iToColor
;				   |								512 = Error setting $iFromIntense
;				   |								1024 = Error setting $iToIntense
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings have been successfully set.
;				   @Error 0 @Extended 0 Return 2 = Success. Transparency Gradient has been successfully turned off.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 11 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Note: Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName, _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapeAreaGradient(ByRef $oDoc, ByRef $oShape, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iBorder = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient
	Local $iError = 0
	Local $avGradient[11]
	Local $sGradName

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$tStyleGradient = $oShape.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iBorder, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LOWriter_ArrayFill($avGradient, $oShape.FillGradientName(), $tStyleGradient.Style(), _
				$oShape.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), ($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	If ($oShape.FillStyle() <> $__LOWCONST_FILL_STYLE_GRADIENT) Then $oShape.FillStyle = $__LOWCONST_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		__LOWriter_GradientPresets($oDoc, $oShape, $tStyleGradient, $sGradientName)
		$iError = ($oShape.FillGradientName() = $sGradientName) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOW_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oShape.FillStyle = $__LOWCONST_FILL_STYLE_OFF
			$oShape.FillGradientName = ""
			Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LOWriter_IntIsBetween($iType, $LOW_GRAD_TYPE_LINEAR, $LOW_GRAD_TYPE_RECT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LOWriter_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oShape.FillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oShape.FillGradientStepCount() = $iIncrement) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LOWriter_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$tStyleGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LOWriter_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tStyleGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LOWriter_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$tStyleGradient.Angle = ($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iBorder <> Null) Then
		If Not __LOWriter_IntIsBetween($iBorder, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$tStyleGradient.Border = $iBorder
	EndIf

	If ($iFromColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iFromColor, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
		$tStyleGradient.StartColor = $iFromColor
	EndIf

	If ($iToColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iToColor, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 11, 0)
		$tStyleGradient.EndColor = $iToColor
	EndIf

	If ($iFromIntense <> Null) Then
		If Not __LOWriter_IntIsBetween($iFromIntense, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 12, 0)
		$tStyleGradient.StartIntensity = $iFromIntense
	EndIf

	If ($iToIntense <> Null) Then
		If Not __LOWriter_IntIsBetween($iToIntense, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 13, 0)
		$tStyleGradient.EndIntensity = $iToIntense
	EndIf

	If ($oShape.FillGradientName() = "") Then

		$sGradName = __LOWriter_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

		$oShape.FillGradientName = $sGradName
		If ($oShape.FillGradientName <> $sGradName) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
	EndIf

	$oShape.FillGradient = $tStyleGradient

	; Error checking
	$iError = ($iType = Null) ? $iError : ($oShape.FillGradient.Style() = $iType) ? $iError : BitOR($iError, 2)
	$iError = ($iXCenter = Null) ? $iError : ($oShape.FillGradient.XOffset() = $iXCenter) ? $iError : BitOR($iError, 8)
	$iError = ($iYCenter = Null) ? $iError : ($oShape.FillGradient.YOffset() = $iYCenter) ? $iError : BitOR($iError, 16)
	$iError = ($iAngle = Null) ? $iError : (($oShape.FillGradient.Angle() / 10) = $iAngle) ? $iError : BitOR($iError, 32)
	$iError = ($iBorder = Null) ? $iError : ($oShape.FillGradient.Border() = $iBorder) ? $iError : BitOR($iError, 64)
	$iError = ($iFromColor = Null) ? $iError : ($oShape.FillGradient.StartColor() = $iFromColor) ? $iError : BitOR($iError, 128)
	$iError = ($iToColor = Null) ? $iError : ($oShape.FillGradient.EndColor() = $iToColor) ? $iError : BitOR($iError, 256)
	$iError = ($iFromIntense = Null) ? $iError : ($oShape.FillGradient.StartIntensity() = $iFromIntense) ? $iError : BitOR($iError, 512)
	$iError = ($iToIntense = Null) ? $iError : ($oShape.FillGradient.EndIntensity() = $iToIntense) ? $iError : BitOR($iError, 1024)

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapeGetAnchor
; Description ...: Create a Text Cursor at the Shape Anchor position.
; Syntax ........: _LOWriter_ShapeGetAnchor(Byref $oShape)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by previous _LOWriter_ShapeInsert, or _LOWriter_ShapeGetObjByName function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Shape anchor Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully returned the Shape Anchor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ShapeInsert, _LOWriter_ShapeGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapeGetAnchor(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnchor

	If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oAnchor = $oShape.Anchor.Text.createTextCursorByRange($oShape.Anchor())
	If Not IsObj($oAnchor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oAnchor)
EndFunc



; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_GetShapeName
; Description ...: Create a Shape Name that hasn't been used yet in the document.
; Syntax ........: __LOWriter_GetShapeName(ByRef $oDoc, $sShapeName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sShapeName          - a string value. The Shape name to begin with.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sShapeName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retreive DrawPage object.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Document contained no shapes, returns the Shape name with a "1" appended.
;				   @Error 0 @Extended 1 Return String = Success. Returns the unique Shape name to use.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function adds a digit after the shape name, incrementing it until that name hasn't been used yet in L.O.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_GetShapeName(ByRef $oDoc, $sShapeName)
Local $oShapes

	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oShapes = $oDoc.DrawPage()
	If Not IsObj($oShapes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $oShapes.hasElements() Then

		For $i = 1 To $oShapes.getCount + 1

		For $j = 0 To $oShapes.getCount() - 1
			If ($oShapes.getByIndex($j).Name() = $sShapeName & $i) Then ExitLoop

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next

If ($oShapes.getByIndex($j).Name() <> $sShapeName & $i) Then ExitLoop;If no matches, exit loop with current name.
		Next

	Else

	Return SetError($__LOW_STATUS_SUCCESS,0,$sShapeName & "1");If Doc has no shapes, just return the name with a "1" appended.
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 1, $sShapeName & $i)
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateArrow
; Description ...: Create a Arrow type Shape.
; Syntax ........: __LOWriter_Shape_CreateArrow($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (0-25). The Type of shape to create. See $LOW_SHAPE_TYPE_ARROWS_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" or "com.sun.star.drawing.EllipseShape" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;				   @Error 2 @Extended 3 Return 0 = Failed to create "MirroredX" property structure.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve the Position Structure.
;				   @Error 2 @Extended 5 Return 0 = Failed to retrieve the Size Structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to create a unique Shape name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;					$LOW_SHAPE_TYPE_ARROWS_ARROW_S_SHAPED $LOW_SHAPE_TYPE_ARROWS_ARROW_SPLIT, $LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT,
;					$LOW_SHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateArrow($oDoc, $iWidth, $iHeight, $iShapeType)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tProp, $tProp2, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

$oShape =  $oDoc.createInstance("com.sun.star.drawing.CustomShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

$tProp = __LOWriter_SetPropertyValue("Type", "")
If @error Then Return SetError($__LOW_STATUS_INIT_ERROR,2, 0)

Switch $iShapeType

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_4_WAY
		$tProp.Value = "quad-arrow"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY
		$tProp.Value = "quad-arrow-callout"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN
		$tProp.Value = "down-arrow-callout"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT
		$tProp.Value = "left-arrow-callout"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT
		$tProp.Value = "left-right-arrow-callout"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT
		$tProp.Value = "right-arrow-callout"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP
		$tProp.Value = "up-arrow-callout"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN
		$tProp.Value = "up-down-arrow-callout"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT
		$tProp.Value = "mso-spt100"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CIRCULAR
		$tProp.Value = "circular-arrow"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT
		$tProp.Value = "corner-right-arrow" ; "non-primitive"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_DOWN
		$tProp.Value = "down-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_LEFT
		$tProp.Value = "left-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT
		$tProp.Value = "left-right-arrow"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT
		$tProp.Value = "notched-right-arrow"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT
		$tProp.Value = "right-arrow"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT
		$tProp.Value = "split-arrow" ; "non-primitive"??

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_S_SHAPED
		$tProp.Value = "s-sharped-arrow"; "non-primitive"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_SPLIT
		$tProp.Value = "split-arrow" ; "non-primitive"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT
		$tProp.Value = "striped-right-arrow"; "mso-spt100"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_UP
		$tProp.Value = "up-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_DOWN
		$tProp.Value = "up-down-arrow"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT
		$tProp.Value = "up-right-arrow-callout"; "mso-spt89"

	Case $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN
		$tProp.Value = "up-right-down-arrow"; "mso-spt100"

$tProp2 = __LOWriter_SetPropertyValue("MirroredX", True);Shape is an up and left arrow without this Property.
If @error Then Return SetError($__LOW_STATUS_INIT_ERROR,3, 0)

ReDim $atCusShapeGeo[2]
$atCusShapeGeo[1] = $tProp2

	Case $LOW_SHAPE_TYPE_ARROWS_CHEVRON
		$tProp.Value = "chevron"

	Case $LOW_SHAPE_TYPE_ARROWS_PENTAGON
		$tProp.Value = "pentagon-right"

	EndSwitch

$atCusShapeGeo[0] = $tProp
$oShape.CustomShapeGeometry = $atCusShapeGeo

$tPos = $oShape.Position()
If Not IsObj($tPos) Then Return SetError($__LOW_STATUS_INIT_ERROR,4,0)

$tPos.X = 0
$tPos.Y = 0

$oShape.Position = $tPos

$tSize = $oShape.Size()
If Not IsObj($tSize) Then Return SetError($__LOW_STATUS_INIT_ERROR,5,0)

$tSize.Width = $iWidth
$tSize.Height = $iHeight

$oShape.Size = $tSize

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

Return SetError($__LOW_STATUS_SUCCESS,0,$oShape)
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateBasic
; Description ...: Create a Basic type Shape.
; Syntax ........: __LOWriter_Shape_CreateBasic($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (26-49). The Type of shape to create. See $LOW_SHAPE_TYPE_BASIC_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" or "com.sun.star.drawing.EllipseShape" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to create a unique Shape name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;					$LOW_SHAPE_TYPE_BASIC_CIRCLE_PIE, $LOW_SHAPE_TYPE_BASIC_FRAME
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateBasic($oDoc, $iWidth, $iHeight, $iShapeType)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]
	Local $iCircleKind_CUT = 2; a circle with a cut connected by a line.
	Local $iCircleKind_ARC = 3; a circle with an open cut.

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

If ($iShapeType = $LOW_SHAPE_TYPE_BASIC_CIRCLE_SEGMENT) Or ($iShapeType = $LOW_SHAPE_TYPE_BASIC_ARC) Then; These two shapes need special procedures.
$oShape =  $oDoc.createInstance("com.sun.star.drawing.EllipseShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

	Else
$oShape =  $oDoc.createInstance("com.sun.star.drawing.CustomShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

$tProp = __LOWriter_SetPropertyValue("Type", "")
If @error Then Return SetError($__LOW_STATUS_INIT_ERROR,2, 0)

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)
EndIf

Switch $iShapeType

	Case $LOW_SHAPE_TYPE_BASIC_ARC
$oShape.FillColor = $LOW_COLOR_OFF

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Elliptical arc ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

	$oShape.CircleKind = $iCircleKind_ARC
	$oShape.CircleStartAngle = 0
	$oShape.CircleEndAngle = 25000

	Case $LOW_SHAPE_TYPE_BASIC_ARC_BLOCK
		$tProp.Value = "block-arc"

	Case $LOW_SHAPE_TYPE_BASIC_CIRCLE_PIE
		$tProp.Value = "circle-pie" ; "mso-spt100"

	Case $LOW_SHAPE_TYPE_BASIC_CIRCLE_SEGMENT
$oShape.Name = __LOWriter_GetShapeName($oDoc, "Ellipse Segment ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

	$oShape.CircleKind = $iCircleKind_CUT
	$oShape.CircleStartAngle = 0
	$oShape.CircleEndAngle = 25000

	Case $LOW_SHAPE_TYPE_BASIC_CROSS
		$tProp.Value = "cross"

	Case $LOW_SHAPE_TYPE_BASIC_CUBE
		$tProp.Value = "cube"

	Case $LOW_SHAPE_TYPE_BASIC_CYLINDER
		$tProp.Value = "can"

	Case $LOW_SHAPE_TYPE_BASIC_DIAMOND
		$tProp.Value = "diamond"

		Case $LOW_SHAPE_TYPE_BASIC_ELLIPSE, $LOW_SHAPE_TYPE_BASIC_CIRCLE
		$tProp.Value = "ellipse"

	Case $LOW_SHAPE_TYPE_BASIC_FOLDED_CORNER
		$tProp.Value = "paper"

	Case $LOW_SHAPE_TYPE_BASIC_FRAME
		$tProp.Value = "frame" ;Not working

	Case $LOW_SHAPE_TYPE_BASIC_HEXAGON
		$tProp.Value = "hexagon"

	Case $LOW_SHAPE_TYPE_BASIC_OCTAGON
		$tProp.Value = "octagon"

		Case $LOW_SHAPE_TYPE_BASIC_PARALLELOGRAM
		$tProp.Value = "parallelogram"

	Case $LOW_SHAPE_TYPE_BASIC_RECTANGLE, $LOW_SHAPE_TYPE_BASIC_SQUARE
		$tProp.Value = "rectangle"

		Case $LOW_SHAPE_TYPE_BASIC_RECTANGLE_ROUNDED, $LOW_SHAPE_TYPE_BASIC_SQUARE_ROUNDED
		$tProp.Value = "round-rectangle"

	Case $LOW_SHAPE_TYPE_BASIC_REGULAR_PENTAGON
		$tProp.Value = "pentagon"

	Case $LOW_SHAPE_TYPE_BASIC_RING
		$tProp.Value = "ring"

		Case $LOW_SHAPE_TYPE_BASIC_TRAPEZOID
		$tProp.Value = "trapezoid"

	Case $LOW_SHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES
		$tProp.Value = "isosceles-triangle"

	Case $LOW_SHAPE_TYPE_BASIC_TRIANGLE_RIGHT
		$tProp.Value = "right-triangle"

	EndSwitch

If ($iShapeType <> $LOW_SHAPE_TYPE_BASIC_CIRCLE_SEGMENT) And ($iShapeType <> $LOW_SHAPE_TYPE_BASIC_ARC) Then
$atCusShapeGeo[0] = $tProp
$oShape.CustomShapeGeometry = $atCusShapeGeo
EndIf

$tPos = $oShape.Position()
If Not IsObj($tPos) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$tPos.X = 0
$tPos.Y = 0

$oShape.Position = $tPos

$tSize = $oShape.Size()
If Not IsObj($tSize) Then Return SetError($__LOW_STATUS_INIT_ERROR,4,0)

$tSize.Width = $iWidth
$tSize.Height = $iHeight

$oShape.Size = $tSize

Return SetError($__LOW_STATUS_SUCCESS,0,$oShape)
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateCallout
; Description ...: Create a Callout type Shape.
; Syntax ........: __LOWriter_Shape_CreateCallout($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (50-56). The Type of shape to create. See $LOW_SHAPE_TYPE_CALLOUT_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to create a unique Shape name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateCallout($oDoc, $iWidth, $iHeight, $iShapeType)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

$oShape =  $oDoc.createInstance("com.sun.star.drawing.CustomShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

$tProp = __LOWriter_SetPropertyValue("Type", "")
If @error Then Return SetError($__LOW_STATUS_INIT_ERROR,2, 0)

Switch $iShapeType

		Case $LOW_SHAPE_TYPE_CALLOUT_CLOUD
		$tProp.Value = "cloud-callout"

		Case $LOW_SHAPE_TYPE_CALLOUT_LINE_1
		$tProp.Value = "line-callout-1"

		Case $LOW_SHAPE_TYPE_CALLOUT_LINE_2
		$tProp.Value = "line-callout-2"

		Case $LOW_SHAPE_TYPE_CALLOUT_LINE_3
		$tProp.Value = "line-callout-3"

	Case $LOW_SHAPE_TYPE_CALLOUT_RECTANGULAR
		$tProp.Value = "rectangular-callout"

		Case $LOW_SHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED
		$tProp.Value = "round-rectangular-callout"

		Case $LOW_SHAPE_TYPE_CALLOUT_ROUND
		$tProp.Value = "round-callout"

	EndSwitch

$atCusShapeGeo[0] = $tProp
$oShape.CustomShapeGeometry = $atCusShapeGeo

$tPos = $oShape.Position()
If Not IsObj($tPos) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$tPos.X = 0
$tPos.Y = 0

$oShape.Position = $tPos

$tSize = $oShape.Size()
If Not IsObj($tSize) Then Return SetError($__LOW_STATUS_INIT_ERROR,4,0)

$tSize.Width = $iWidth
$tSize.Height = $iHeight

$oShape.Size = $tSize

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

Return SetError($__LOW_STATUS_SUCCESS,0,$oShape)
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateFlowchart
; Description ...: Create a FlowChart type Shape.
; Syntax ........: __LOWriter_Shape_CreateFlowchart($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (57-84). The Type of shape to create. See $LOW_SHAPE_TYPE_FLOWCHART_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to create a unique Shape name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateFlowchart($oDoc, $iWidth, $iHeight, $iShapeType)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oShape
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

$oShape =  $oDoc.createInstance("com.sun.star.drawing.CustomShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

$tProp = __LOWriter_SetPropertyValue("Type", "")
If @error Then Return SetError($__LOW_STATUS_INIT_ERROR,2, 0)

Switch $iShapeType

	Case $LOW_SHAPE_TYPE_FLOWCHART_CARD
		$tProp.Value = "flowchart-card"

	Case $LOW_SHAPE_TYPE_FLOWCHART_COLLATE
		$tProp.Value = "flowchart-collate"

	Case $LOW_SHAPE_TYPE_FLOWCHART_CONNECTOR
		$tProp.Value = "flowchart-connector"

	Case $LOW_SHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE
		$tProp.Value = "flowchart-off-page-connector"

		Case $LOW_SHAPE_TYPE_FLOWCHART_DATA
		$tProp.Value = "flowchart-data"

		Case $LOW_SHAPE_TYPE_FLOWCHART_DECISION
		$tProp.Value = "flowchart-decision"

	Case $LOW_SHAPE_TYPE_FLOWCHART_DELAY
		$tProp.Value = "flowchart-delay"

	Case $LOW_SHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE
		$tProp.Value = "flowchart-direct-access-storage"

	Case $LOW_SHAPE_TYPE_FLOWCHART_DISPLAY
		$tProp.Value = "flowchart-display"

		Case $LOW_SHAPE_TYPE_FLOWCHART_DOCUMENT
		$tProp.Value = "flowchart-document"

	Case $LOW_SHAPE_TYPE_FLOWCHART_EXTRACT
		$tProp.Value = "flowchart-extract"

		Case $LOW_SHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE
		$tProp.Value = "flowchart-internal-storage"

	Case $LOW_SHAPE_TYPE_FLOWCHART_MAGNETIC_DISC
		$tProp.Value = "flowchart-magnetic-disk"

	Case $LOW_SHAPE_TYPE_FLOWCHART_MANUAL_INPUT
		$tProp.Value = "flowchart-manual-input"

	Case $LOW_SHAPE_TYPE_FLOWCHART_MANUAL_OPERATION
		$tProp.Value = "flowchart-manual-operation"

	Case $LOW_SHAPE_TYPE_FLOWCHART_MERGE
		$tProp.Value = "flowchart-merge"

	Case $LOW_SHAPE_TYPE_FLOWCHART_MULTIDOCUMENT
		$tProp.Value = "flowchart-multidocument"

	Case $LOW_SHAPE_TYPE_FLOWCHART_OR
		$tProp.Value = "flowchart-or"

	Case $LOW_SHAPE_TYPE_FLOWCHART_PREPARATION
		$tProp.Value = "flowchart-preparation"

	Case $LOW_SHAPE_TYPE_FLOWCHART_PROCESS
		$tProp.Value = "flowchart-process"

		Case $LOW_SHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE
		$tProp.Value = "flowchart-alternate-process"

		Case $LOW_SHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED
		$tProp.Value = "flowchart-predefined-process"

	Case $LOW_SHAPE_TYPE_FLOWCHART_PUNCHED_TAPE
		$tProp.Value = "flowchart-punched-tape"

	Case $LOW_SHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS
		$tProp.Value = "flowchart-sequential-access"

	Case $LOW_SHAPE_TYPE_FLOWCHART_SORT
		$tProp.Value = "flowchart-sort"

	Case $LOW_SHAPE_TYPE_FLOWCHART_STORED_DATA
		$tProp.Value = "flowchart-stored-data"

	Case $LOW_SHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION
		$tProp.Value = "flowchart-summing-junction"

	Case $LOW_SHAPE_TYPE_FLOWCHART_TERMINATOR
		$tProp.Value = "flowchart-terminator"

	EndSwitch

$atCusShapeGeo[0] = $tProp
$oShape.CustomShapeGeometry = $atCusShapeGeo

$tPos = $oShape.Position()
If Not IsObj($tPos) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$tPos.X = 0
$tPos.Y = 0

$oShape.Position = $tPos

$tSize = $oShape.Size()
If Not IsObj($tSize) Then Return SetError($__LOW_STATUS_INIT_ERROR,4,0)

$tSize.Width = $iWidth
$tSize.Height = $iHeight

$oShape.Size = $tSize

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

Return SetError($__LOW_STATUS_SUCCESS,0,$oShape)
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateLine
; Description ...: Create a Line type Shape.
; Syntax ........: __LOWriter_Shape_CreateLine($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (85-92). The Type of shape to create. See $LOW_SHAPE_TYPE_LINE_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.PolyPolygonBezierCoords" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create the requested Line type Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to create a Position structure.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve the Position Structure.
;				   @Error 2 @Extended 5 Return 0 = Failed to retrieve the Size Structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to create a unique Shape name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateLine($oDoc, $iWidth, $iHeight, $iShapeType)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tSize, $tPolyCoords, $tPos
Local $atPoint[0], $aiFlag[0], $avArray[1]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

$tPolyCoords = __LOWriter_CreateStruct("com.sun.star.drawing.PolyPolygonBezierCoords")
If @error Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

Switch $iShapeType

		Case $LOW_SHAPE_TYPE_LINE_CURVE
$oShape =  $oDoc.createInstance("com.sun.star.drawing.OpenBezierShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

ReDim $atPoint[4]
ReDim $aiFlag[4]

$atPoint[0] = __LOWriter_CreatePoint(0, 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth / 3), Int($iHeight / 2))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth / 2), Int($iHeight))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[3] = __LOWriter_CreatePoint(Int($iWidth), 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$aiFlag[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[1] = $LOW_SHAPE_POINT_TYPE_CONTROL
$aiFlag[2] = $LOW_SHAPE_POINT_TYPE_CONTROL
$aiFlag[3] = $LOW_SHAPE_POINT_TYPE_NORMAL

$avArray[0] = $atPoint
$tPolyCoords.Coordinates = $avArray

$avArray[0] = $aiFlag
$tPolyCoords.Flags = $avArray

$oShape.PolyPolygonBezier = $tPolyCoords

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Bézier curve ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

$oShape.FillColor = $LOW_COLOR_OFF

		Case $LOW_SHAPE_TYPE_LINE_CURVE_FILLED
$oShape =  $oDoc.createInstance("com.sun.star.drawing.ClosedBezierShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

ReDim $atPoint[4]
ReDim $aiFlag[4]

$atPoint[0] = __LOWriter_CreatePoint(0, 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth / 3), Int($iHeight / 2))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth / 2), Int($iHeight))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[3] = __LOWriter_CreatePoint(Int($iWidth), 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$aiFlag[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[1] = $LOW_SHAPE_POINT_TYPE_CONTROL
$aiFlag[2] = $LOW_SHAPE_POINT_TYPE_CONTROL
$aiFlag[3] = $LOW_SHAPE_POINT_TYPE_NORMAL

$avArray[0] = $atPoint
$tPolyCoords.Coordinates = $avArray

$avArray[0] = $aiFlag
$tPolyCoords.Flags = $avArray

$oShape.PolyPolygonBezier = $tPolyCoords

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Bézier curve ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

$oShape.FillColor = 7512015 ; Light blue

		Case $LOW_SHAPE_TYPE_LINE_FREEFORM_LINE
$oShape =  $oDoc.createInstance("com.sun.star.drawing.OpenBezierShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

ReDim $atPoint[3]
ReDim $aiFlag[3]

$atPoint[0] = __LOWriter_CreatePoint(0, 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth / 2), Int($iHeight / 2))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth), Int($iHeight))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

$aiFlag[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[1] = $LOW_SHAPE_POINT_TYPE_CONTROL
$aiFlag[2] = $LOW_SHAPE_POINT_TYPE_NORMAL

$avArray[0] = $atPoint
$tPolyCoords.Coordinates = $avArray

$avArray[0] = $aiFlag
$tPolyCoords.Flags = $avArray

$oShape.PolyPolygonBezier = $tPolyCoords

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Bézier curve ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

		Case $LOW_SHAPE_TYPE_LINE_FREEFORM_LINE_FILLED
$oShape =  $oDoc.createInstance("com.sun.star.drawing.ClosedBezierShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

ReDim $atPoint[5]
ReDim $aiFlag[5]

$atPoint[0] = __LOWriter_CreatePoint(0, 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth / 2), Int($iHeight / 2))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth), Int($iHeight))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[3] = __LOWriter_CreatePoint(0, Int($iHeight))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[4] = __LOWriter_CreatePoint(0, 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$aiFlag[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[1] = $LOW_SHAPE_POINT_TYPE_CONTROL
$aiFlag[2] = $LOW_SHAPE_POINT_TYPE_CONTROL
$aiFlag[3] = $LOW_SHAPE_POINT_TYPE_CONTROL
$aiFlag[4] = $LOW_SHAPE_POINT_TYPE_NORMAL

$avArray[0] = $atPoint
$tPolyCoords.Coordinates = $avArray

$avArray[0] = $aiFlag
$tPolyCoords.Flags = $avArray

$oShape.PolyPolygonBezier = $tPolyCoords

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Bézier curve ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

$oShape.FillColor = 7512015 ; Light blue

	Case $LOW_SHAPE_TYPE_LINE_LINE
$oShape =  $oDoc.createInstance("com.sun.star.drawing.LineShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

ReDim $atPoint[2]
ReDim $aiFlag[2]

$atPoint[0] = __LOWriter_CreatePoint(0, 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth), Int($iHeight))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

$aiFlag[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[1] = $LOW_SHAPE_POINT_TYPE_NORMAL

$avArray[0] = $atPoint
$tPolyCoords.Coordinates = $avArray

$avArray[0] = $aiFlag
$tPolyCoords.Flags = $avArray

$oShape.PolyPolygonBezier = $tPolyCoords

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Line ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

		Case $LOW_SHAPE_TYPE_LINE_POLYGON, $LOW_SHAPE_TYPE_LINE_POLYGON_45
$oShape =  $oDoc.createInstance("com.sun.star.drawing.PolyPolygonShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

ReDim $atPoint[5]
ReDim $aiFlag[5]

$atPoint[0] = __LOWriter_CreatePoint(0, 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth), 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth), Int($iHeight))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[3] = __LOWriter_CreatePoint(0, Int($iHeight))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[4] = __LOWriter_CreatePoint(0, 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$aiFlag[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[1] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[2] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[3] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[4] = $LOW_SHAPE_POINT_TYPE_NORMAL

$avArray[0] = $atPoint
$tPolyCoords.Coordinates = $avArray

$avArray[0] = $aiFlag
$tPolyCoords.Flags = $avArray

$oShape.PolyPolygonBezier = $tPolyCoords

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Polygon 4 corners ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

$oShape.FillColor = $LOW_COLOR_OFF

	Case $LOW_SHAPE_TYPE_LINE_POLYGON_45_FILLED
$oShape =  $oDoc.createInstance("com.sun.star.drawing.PolyPolygonShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,2,0)

ReDim $atPoint[5]
ReDim $aiFlag[5]

$atPoint[0] = __LOWriter_CreatePoint(0, 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth), 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth), Int($iHeight))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[3] = __LOWriter_CreatePoint(0, Int($iHeight))
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$atPoint[4] = __LOWriter_CreatePoint(0, 0)
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$aiFlag[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[1] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[2] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[3] = $LOW_SHAPE_POINT_TYPE_NORMAL
$aiFlag[4] = $LOW_SHAPE_POINT_TYPE_NORMAL

$avArray[0] = $atPoint
$tPolyCoords.Coordinates = $avArray

$avArray[0] = $aiFlag
$tPolyCoords.Flags = $avArray

$oShape.PolyPolygonBezier = $tPolyCoords

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Polygon 4 corners ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

$oShape.FillColor = 7512015 ; Light blue

	EndSwitch

$tSize = $oShape.Size()
If Not IsObj($tSize) Then Return SetError($__LOW_STATUS_INIT_ERROR,4,0)

$tSize.Width = $iWidth
$tSize.Height = $iHeight

$oShape.Size = $tSize

$tPos = $oShape.Position()
If Not IsObj($tPos) Then Return SetError($__LOW_STATUS_INIT_ERROR,5,0)

$tPos.X = 0
$tPos.Y = 0

$oShape.Position = $tPos

Return SetError($__LOW_STATUS_SUCCESS,0,$oShape)
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateStars
; Description ...: Create a Star or Banner type Shape.
; Syntax ........: __LOWriter_Shape_CreateStars($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (93-104). The Type of shape to create. See $LOW_SHAPE_TYPE_STARS_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to create a unique Shape name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;					$LOW_SHAPE_TYPE_STARS_6_POINT, $LOW_SHAPE_TYPE_STARS_12_POINT, $LOW_SHAPE_TYPE_STARS_SIGNET, $LOW_SHAPE_TYPE_STARS_6_POINT_CONCAVE.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateStars($oDoc, $iWidth, $iHeight, $iShapeType)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oShape
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

$oShape =  $oDoc.createInstance("com.sun.star.drawing.CustomShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

$tProp = __LOWriter_SetPropertyValue("Type", "")
If @error Then Return SetError($__LOW_STATUS_INIT_ERROR,2, 0)

Switch $iShapeType

	Case $LOW_SHAPE_TYPE_STARS_4_POINT
		$tProp.Value = "star4"

		Case $LOW_SHAPE_TYPE_STARS_5_POINT
		$tProp.Value = "star5"

		Case $LOW_SHAPE_TYPE_STARS_6_POINT
		$tProp.Value = "star6"; "non-primitive"

	Case $LOW_SHAPE_TYPE_STARS_6_POINT_CONCAVE
		$tProp.Value = "concave-star6"; "non-primitive"

		Case $LOW_SHAPE_TYPE_STARS_8_POINT
		$tProp.Value = "star8"

		Case $LOW_SHAPE_TYPE_STARS_12_POINT
		$tProp.Value = "star12"; "non-primitive"

		Case $LOW_SHAPE_TYPE_STARS_24_POINT
		$tProp.Value = "star24"

	Case $LOW_SHAPE_TYPE_STARS_DOORPLATE
		$tProp.Value = "mso-spt21"; "doorplate"

		Case $LOW_SHAPE_TYPE_STARS_EXPLOSION
		$tProp.Value = "bang"

	Case $LOW_SHAPE_TYPE_STARS_SCROLL_HORIZONTAL
		$tProp.Value = "horizontal-scroll"

	Case $LOW_SHAPE_TYPE_STARS_SCROLL_VERTICAL
		$tProp.Value = "vertical-scroll"

	Case $LOW_SHAPE_TYPE_STARS_SIGNET
		$tProp.Value = "signet"; "non-primitive"

EndSwitch

$atCusShapeGeo[0] = $tProp
$oShape.CustomShapeGeometry = $atCusShapeGeo

$tPos = $oShape.Position()
If Not IsObj($tPos) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$tPos.X = 0
$tPos.Y = 0

$oShape.Position = $tPos

$tSize = $oShape.Size()
If Not IsObj($tSize) Then Return SetError($__LOW_STATUS_INIT_ERROR,4,0)

$tSize.Width = $iWidth
$tSize.Height = $iHeight

$oShape.Size = $tSize

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

Return SetError($__LOW_STATUS_SUCCESS,0,$oShape)
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateSymbol
; Description ...: Create a Symbol type Shape.
; Syntax ........: __LOWriter_Shape_CreateSymbol($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (105-122). The Type of shape to create. See $LOW_SHAPE_TYPE_SYMBOL_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve the Position Structure.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve the Size Structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to create a unique Shape name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;					$LOW_SHAPE_TYPE_SYMBOL_CLOUD, $LOW_SHAPE_TYPE_SYMBOL_FLOWER, $LOW_SHAPE_TYPE_SYMBOL_PUZZLE, $LOW_SHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, $LOW_SHAPE_TYPE_SYMBOL_BEVEL_DIAMOND
;				   The following shape is visually different from the manually inserted one in L.O. 7.3.4.2:
;					$LOW_SHAPE_TYPE_SYMBOL_LIGHTNING
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateSymbol($oDoc, $iWidth, $iHeight, $iShapeType)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

$oShape =  $oDoc.createInstance("com.sun.star.drawing.CustomShape")
If Not IsObj($oShape) Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

$tProp = __LOWriter_SetPropertyValue("Type", "")
If @error Then Return SetError($__LOW_STATUS_INIT_ERROR,2, 0)

Switch $iShapeType

	Case $LOW_SHAPE_TYPE_SYMBOL_BEVEL_DIAMOND
$tProp.Value = "col-502ad400"

	Case $LOW_SHAPE_TYPE_SYMBOL_BEVEL_OCTAGON
$tProp.Value = "col-60da8460"

	Case $LOW_SHAPE_TYPE_SYMBOL_BEVEL_SQUARE
$tProp.Value = "quad-bevel"

	Case $LOW_SHAPE_TYPE_SYMBOL_BRACE_DOUBLE
$tProp.Value = "brace-pair"
$oShape.FillColor = $LOW_COLOR_OFF

	Case $LOW_SHAPE_TYPE_SYMBOL_BRACE_LEFT
$tProp.Value = "left-brace"
$oShape.FillColor = $LOW_COLOR_OFF

	Case $LOW_SHAPE_TYPE_SYMBOL_BRACE_RIGHT
$tProp.Value = "right-brace"
$oShape.FillColor = $LOW_COLOR_OFF

	Case $LOW_SHAPE_TYPE_SYMBOL_BRACKET_DOUBLE
$tProp.Value = "bracket-pair"
$oShape.FillColor = $LOW_COLOR_OFF

	Case $LOW_SHAPE_TYPE_SYMBOL_BRACKET_LEFT
$tProp.Value = "left-bracket"
$oShape.FillColor = $LOW_COLOR_OFF

	Case $LOW_SHAPE_TYPE_SYMBOL_BRACKET_RIGHT
$tProp.Value = "right-bracket"
$oShape.FillColor = $LOW_COLOR_OFF

		Case $LOW_SHAPE_TYPE_SYMBOL_CLOUD
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "cloud"
$tProp.Value = "cloud"

		Case $LOW_SHAPE_TYPE_SYMBOL_FLOWER
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "flower"
$tProp.Value = "flower"

		Case $LOW_SHAPE_TYPE_SYMBOL_HEART
$tProp.Value = "heart"

		Case $LOW_SHAPE_TYPE_SYMBOL_LIGHTNING
;~ Custom Shape Geometry Type = "non-primitive" ???? Try "lightning"
$tProp.Value = "lightning"

		Case $LOW_SHAPE_TYPE_SYMBOL_MOON
$tProp.Value = "moon"

	Case $LOW_SHAPE_TYPE_SYMBOL_SMILEY
$tProp.Value = "smiley"

		Case $LOW_SHAPE_TYPE_SYMBOL_SUN
$tProp.Value = "sun"

	Case $LOW_SHAPE_TYPE_SYMBOL_PROHIBITED
$tProp.Value = "forbidden"

	Case $LOW_SHAPE_TYPE_SYMBOL_PUZZLE
$tProp.Value = "puzzle"

	EndSwitch

$atCusShapeGeo[0] = $tProp
$oShape.CustomShapeGeometry = $atCusShapeGeo

$tPos = $oShape.Position()
If Not IsObj($tPos) Then Return SetError($__LOW_STATUS_INIT_ERROR,3,0)

$tPos.X = 0
$tPos.Y = 0

$oShape.Position = $tPos

$tSize = $oShape.Size()
If Not IsObj($tSize) Then Return SetError($__LOW_STATUS_INIT_ERROR,4,0)

$tSize.Width = $iWidth
$tSize.Height = $iHeight

$oShape.Size = $tSize

$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR,1,0)

Return SetError($__LOW_STATUS_SUCCESS,0,$oShape)
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CreatePoint
; Description ...: Creates a Position structure.
; Syntax ........: __LOWriter_CreatePoint($iX, $iY)
; Parameters ....: $iX                  - an integer value. The X position, in Micrometers.
;                  $iY                  - an integer value. The Y position, in Micrometers.
; Return values .: Success: Structure
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $iX not an Integer.
;				   @Error 1 @Extended 2 Return 0 = $iY not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create a Position Structure.
;				   --Success--
;				   @Error 0 @Extended 0 Return Structure = Success. Returning created Position Structure set to $iX and $iY values.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Modified from A. Pitonyak, Listing 493. in OOME 3.0
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CreatePoint($iX, $iY)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $tPoint

If Not IsInt($iX) Then Return SetError($__LOW_STATUS_INPUT_ERROR,1,0)
If Not IsInt($iY) Then Return SetError($__LOW_STATUS_INPUT_ERROR,2,0)

$tPoint = __LOWriter_CreateStruct("com.sun.star.awt.Point")
If @error Then Return SetError($__LOW_STATUS_INIT_ERROR,1,0)

$tPoint.X = $iX
$tPoint.Y = $iY

Return SetError($__LOW_STATUS_SUCCESS, 0, $tPoint)
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ShapeArrowStyleName
; Description ...: Convert a Arrow head Constant to the corresponding name or reverse.
; Syntax ........: __LOWriter_ShapeArrowStyleName([$iArrowStyle = Null[, $sArrowStyle = Null]])
; Parameters ....: $iArrowStyle         - [optional] an integer value (0-32). Default is Null. The Arrow Style Constant to convert to its corresponding name. See $LOW_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeWriter_Constants.au3
;                  $sArrowStyle         - [optional] a string value. Default is Null. The Arrow Style Name to convert to the corresponding constant if found.
; Return values .: Success: String or Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $iArrowStyle not set to Null, not an Integer, less than 0, or greater than Arrow type constants. See $LOW_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeWriter_Constants.au3
;				   @Error 1 @Extended 2 Return 0 = $sArrowStyle not a String and not set to Null.
;				   @Error 1 @Extended 3 Return 0 = Both $iArrowStyle and $sArrowStyle set to Null.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Constant called in $iArrowStyle was successfully converted to its corresponding Arrow Type Name.
;				   @Error 0 @Extended 1 Return Integer = Success. Arrow Type Name called in $sArrowStyle was successfully converted to its corresponding Constant value.
;				   @Error 0 @Extended 2 Return String = Success. Arrow Type Name called in $sArrowStyle was not matched to an existing Constant value, returning called name. Possibly a custom value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ShapeArrowStyleName($iArrowStyle = Null, $sArrowStyle = Null)
Local $asArrowStyles[33]

		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_NONE] = ""
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_ARROW_SHORT] = "Arrow short"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CONCAVE_SHORT] = "Concave short"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_ARROW] = "Arrow"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_TRIANGLE] = "Triangle"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CONCAVE] = "Concave"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_ARROW_LARGE] = "Arrow large"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CIRCLE] = "Circle"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE] = "Square"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_45] = "Square 45"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DIAMOND] = "Diamond"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE] = "Half Circle"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINES] = "Dimension Lines"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINE_ARROW] = "Dimension Line Arrow"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DIMENSION_LINE] = "Dimension Line"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_LINE_SHORT] = "Line short"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_LINE] = "Line"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_TRIANGLE_UNFILLED] = "Triangle unfilled"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DIAMOND_UNFILLED] = "Diamond unfilled"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CIRCLE_UNFILLED] = "Circle unfilled"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED] = "Square 45 unfilled"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_UNFILLED] = "Square unfilled"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE_UNFILLED] = "Half Circle unfilled"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_LEFT] = "Half Arrow left"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_RIGHT] = "Half Arrow right"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_REVERSED_ARROW] = "Reversed Arrow"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DOUBLE_ARROW] = "Double Arrow"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_ONE] = "CF One"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_ONLY_ONE] = "CF Only One"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_MANY] = "CF Many"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_MANY_ONE] = "CF Many One"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_ZERO_ONE] = "CF Zero One"
		$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY] = "CF Zero Many"

If ($iArrowStyle <> Null) Then
If Not __LOWriter_IntIsBetween($iArrowStyle, 0, UBound($asArrowStyles) -1) Then Return SetError($__LOW_STATUS_INPUT_ERROR,1,0)

Return SetError($__LOW_STATUS_SUCCESS,0,$asArrowStyles[$iArrowStyle]); Return the requested Arrow Style name.

ElseIf ($sArrowStyle <> Null) Then
If Not IsString($sArrowStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR,2,0)


For $i = 0 To UBound($asArrowStyles) -1

If ($asArrowStyles[$i] = $sArrowStyle) Then Return SetError($__LOW_STATUS_SUCCESS,1,$i); Return the array element where the matching Arrow Style was found.

Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
Next

Return SetError($__LOW_STATUS_SUCCESS,2,$sArrowStyle); If no matches, just return the name, as it could be a custom value.

Else
	Return SetError($__LOW_STATUS_INPUT_ERROR,3,0); No vaues called.

	EndIf
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ShapeLineStyleName
; Description ...:  Convert a Line Style Constant to the corresponding name or reverse.
; Syntax ........: __LOWriter_ShapeLineStyleName([$iLineStyle = Null[, $sLineStyle = Null]])
; Parameters ....: $iLineStyle          - [optional] an integer value. Default is Null. The Line Style Constant to convert to its corresponding name. See $LOW_SHAPE_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3
;                  $sLineStyle          - [optional] a string value. Default is Null. The Line Style Name to convert to the corresponding constant if found.
; Return values .: Success: String or Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $iLineStyle not set to Null, not an Integer, less than 0, or greater than Line Style constants. See $LOW_SHAPE_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3
;				   @Error 1 @Extended 2 Return 0 = $sLineStyle not a String and not set to Null.
;				   @Error 1 @Extended 3 Return 0 = Both $iLineStyle and $sLineStyle set to Null.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Constant called in $iLineStyle was successfully converted to its corresponding Line Style Name.
;				   @Error 0 @Extended 1 Return Integer = Success. Line Style Name called in $sLineStyle was successfully converted to its corresponding Constant value.
;				   @Error 0 @Extended 2 Return String = Success. Line Style Name called in $sLineStyle was not matched to an existing Constant value, returning called name. Possibly a custom value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ShapeLineStyleName($iLineStyle = Null, $sLineStyle = Null)
Local $asLineStyles[32]

; $LOW_SHAPE_LINE_STYLE_NONE, $LOW_SHAPE_LINE_STYLE_CONTINUOUS, don't have a name, so to keep things symmetrical I created my own, but those two won't be used.
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_NONE] = "NONE"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_CONTINUOUS] = "CONTINUOUS"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOT] = "Dot"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOT_ROUNDED] = "Dot (Rounded)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DOT] = "Long Dot"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DOT_ROUNDED] = "Long Dot (Rounded)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH] = "Dash"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH_ROUNDED] = "Dash (Rounded)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DASH] = "Long Dash"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DASH_ROUNDED] = "Long Dash (Rounded)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH] = "Double Dash"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_ROUNDED] = "Double Dash (Rounded)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH_DOT] = "Dash Dot"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH_DOT_ROUNDED] = "Dash Dot (Rounded)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DASH_DOT] = "Long Dash Dot"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DASH_DOT_ROUNDED] = "Long Dash Dot (Rounded)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT] = "Double Dash Dot"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_ROUNDED] = "Double Dash Dot (Rounded)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH_DOT_DOT] = "Dash Dot Dot"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH_DOT_DOT_ROUNDED] = "Dash Dot Dot (Rounded)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT] = "Double Dash Dot Dot"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT_ROUNDED] = "Double Dash Dot Dot (Rounded)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_ULTRAFINE_DOTTED] = "Ultrafine Dotted (var)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_FINE_DOTTED] = "Fine Dotted"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_ULTRAFINE_DASHED] = "Ultrafine Dashed"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_FINE_DASHED] = "Fine Dashed"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASHED] = "Dashed (var)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_LINE_STYLE_9] = "Line Style 9"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_3_DASHES_3_DOTS] = "3 Dashes 3 Dots (var)"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_ULTRAFINE_2_DOTS_3_DASHES] = "Ultrafine 2 Dots 3 Dashes"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_2_DOTS_1_DASH] = "2 Dots 1 Dash"
		$asLineStyles[$LOW_SHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS] = "Line with Fine Dots"

If ($iLineStyle <> Null) Then
If Not __LOWriter_IntIsBetween($iLineStyle, 0, UBound($asLineStyles) -1) Then Return SetError($__LOW_STATUS_INPUT_ERROR,1,0)

Return SetError($__LOW_STATUS_SUCCESS,0,$asLineStyles[$iLineStyle]); Return the requested Line Style name.

ElseIf ($sLineStyle <> Null) Then
If Not IsString($sLineStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR,2,0)


For $i = 0 To UBound($asLineStyles) -1

If ($asLineStyles[$i] = $sLineStyle) Then Return SetError($__LOW_STATUS_SUCCESS,1,$i); Return the array element where the matching Line Style was found.

Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
Next

Return SetError($__LOW_STATUS_SUCCESS,2,$sLineStyle); If no matches, just return the name, as it could be a custom value.

Else
	Return SetError($__LOW_STATUS_INPUT_ERROR,3,0); No values called.

	EndIf
EndFunc

