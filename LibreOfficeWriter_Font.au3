;~ #AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter)
; AutoIt Version : v3.3.16.1
; UDF Version    : 0.0.0.3
; Description ...: Provides basic functionality through Autoit for interacting with Libre Office Writer.
; Author(s) .....: donnyh13
; Sources . . . .:  jguinch -- Printmgr.au3, used (_PrintMgr_EnumPrinter);
;					mLipok -- OOoCalc.au3, used (__OOoCalc_ComErrorHandler_UserFunction,_InternalComErrorHandler,
;						-- WriterDemo.au3, used _CreateStruct;
;					Andrew Pitonyak & Laurent Godard (VersionGet);
;					Leagnus & GMK -- OOoCalc.au3, used (SetPropertyValue)
; Dll ...........:
; Note...........: Tips/templates taken from OOoCalc UDF written by user GMK; also from Word UDF by user water.
;					I found the book by Andrew Pitonyak very helpful also, titled, "OpenOffice.org Macros Explained;
;						OOME Third Edition".
;					Of course, this UDF is written using the English version of LibreOffice, and may only work for the English
;						version of LibreOffice installations. Many functions in this UDF may or may not work with OpenOffice
;						Writer, however some settings are definitely for LibreOffice only.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
;_LOWriter_FontExists
;_LOWriter_FontsList
; ===============================================================================================================================

Func _LOWriter_FontExists(ByRef $oDoc, $sFontName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFontName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	For $i = 0 To UBound($atFonts) - 1
		If $atFonts[$i].Name = $sFontName Then Return SetError($__LOW_STATUS_SUCCESS, 0, True)
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next
	Return SetError($__LOW_STATUS_SUCCESS, 0, False)

EndFunc   ;==>_LOWriter_FontExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FontsList
; Description ...: Retrieve a list of currently available fonts.
; Syntax ........: _LOWriter_FontsList(Byref $oDoc)
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous DocOpen, DocConnect, or
;				   +					DocCreate function.
; Return values .:  Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Font list.
;				   --Success--
;				   @Error 0 @Extended ? Return Array  = Success. Returns a 4 Column Array, @extended is set to the number of
;				   +			results. The First column (Array[1][0]) contains the Font Name. The Second column (Array [1][1]
;				   +			contains the style name (Such as Bold Italic etc.) The third column (Array[1][2]) contains
;				   +			the Font weight (Bold ) See Constants listed below; The fourth column (Array[1][3]) Contains
;				   +			the font slant (Italic) See constants below.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Many fonts will be listed multiple times, this is because of the varying settings for them, such as bold,
;					Italic, etc. Style Name is really a repeat of weight(Bold) and Slant (Italic) settings, but is included for
;					easier processing if required. From personal tests, Slant only returns 0 or 2. This function may cause a
;					 processor usage spike for a moment or two. If you wish to eliminate this, comment out the current sleep
;					function and place a sleep(10) in its place.
; Weight Constants : $LOW_WEIGHT_DONT_KNOW(0); The font weight is not specified/known.
;						$LOW_WEIGHT_THIN(50); specifies a 50% font weight.
;						$LOW_WEIGHT_ULTRA_LIGHT(60); specifies a 60% font weight.
;						$LOW_WEIGHT_LIGHT(75); specifies a 75% font weight.
;						$LOW_WEIGHT_SEMI_LIGHT(90); specifies a 90% font weight.
;						$LOW_WEIGHT_NORMAL(100); specifies a normal font weight.
;						$LOW_WEIGHT_SEMI_BOLD(110); specifies a 110% font weight.
;						$LOW_WEIGHT_BOLD(150); specifies a 150% font weight.
;						$LOW_WEIGHT_ULTRA_BOLD(175); specifies a 175% font weight.
;						$LOW_WEIGHT_BLACK(200); specifies a 200% font weight.
; Slant/Posture Constants : $LOW_POSTURE_NONE(0); specifies a font without slant.
;							$LOW_POSTURE_OBLIQUE(1); specifies an oblique font (slant not designed into the font).
;							$LOW_POSTURE_ITALIC(2); specifies an italic font (slant designed into the font).
;							$LOW_POSTURE_DontKnow(3); specifies a font with an unknown slant.
;							$LOW_POSTURE_REV_OBLIQUE(4); specifies a reverse oblique font (slant not designed into the font).
;							$LOW_POSTURE_REV_ITALIC(5); specifies a reverse italic font (slant designed into the font).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FontsList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts
	Local $asFonts[0][4]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	ReDim $asFonts[UBound($atFonts)][4]
;~ $asFonts[0][0] = UBound($atFonts)
	For $i = 0 To UBound($atFonts) - 1
		$asFonts[$i][0] = $atFonts[$i].Name()
		$asFonts[$i][1] = $atFonts[$i].StyleName()
		$asFonts[$i][2] = $atFonts[$i].Weight
		$asFonts[$i][3] = $atFonts[$i].Slant() ;only 0 or 2?
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next

	Return SetError($__LOW_STATUS_SUCCESS, UBound($atFonts), $asFonts)
EndFunc   ;==>_LOWriter_FontsList
