#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"

Example()

Func Example()
	Local $iInch_To_100thMM, $iCM_To_100thMM, $iMM_To_100thMM, $iPt_To_100thMM
	Local $iInch_From_100thMM, $iCM_From_100thMM, $iMM_From_100thMM, $iPt_From_100thMM

	; Convert 1 Inch to Hundredth Millimeters.
	$iInch_To_100thMM = _LO_UnitConvert(1, $LO_CONVERT_UNIT_INCH_100THMM)
	If @error Then _ERROR("Failed to convert to 100th MM from Inch. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 2.54 Centimeters to Hundredth Millimeters.
	$iCM_To_100thMM = _LO_UnitConvert(2.54, $LO_CONVERT_UNIT_CM_100THMM)
	If @error Then _ERROR("Failed to convert to 100th MM from Centimeter. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 25.40 Millimeters to Hundredth Millimeters.
	$iMM_To_100thMM = _LO_UnitConvert(25.4, $LO_CONVERT_UNIT_MM_100THMM)
	If @error Then _ERROR("Failed to convert to 100th MM from Millimeter. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 72 Printer's Points to Hundredth Millimeters.
	$iPt_To_100thMM = _LO_UnitConvert(72, $LO_CONVERT_UNIT_PT_100THMM)
	If @error Then _ERROR("Failed to convert to 100th MM from Printer's Points. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "1 Inch converted to Hundredths of a Millimeter (HMM) = " & $iInch_To_100thMM & @CRLF & _
			"2.54 CM to Hundredths of a Millimeter (HMM) = " & $iCM_To_100thMM & @CRLF & _
			"25.40 MM to Hundredths of a Millimeter (HMM) = " & $iMM_To_100thMM & @CRLF & _
			"72 Printer's Points to Hundredths of a Millimeter (HMM) = " & $iPt_To_100thMM)

	; Convert 2540 Hundredth Millimeters to Inches.
	$iInch_From_100thMM = _LO_UnitConvert(2540, $LO_CONVERT_UNIT_100THMM_INCH)
	If @error Then _ERROR("Failed to convert from 100th MM to Inch. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 2540 Hundredth Millimeters to Centimeters.
	$iCM_From_100thMM = _LO_UnitConvert(2540, $LO_CONVERT_UNIT_100THMM_CM)
	If @error Then _ERROR("Failed to convert from 100th MM to Centimeter. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 2540 Hundredth Millimeters to Millimeters.
	$iMM_From_100thMM = _LO_UnitConvert(2540, $LO_CONVERT_UNIT_100THMM_MM)
	If @error Then _ERROR("Failed to convert from 100th MM to Millimeter. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 2540 Hundredth Millimeters to Printer's Points.
	$iPt_From_100thMM = _LO_UnitConvert(2540, $LO_CONVERT_UNIT_100THMM_PT)
	If @error Then _ERROR("Failed to convert from 100th MM to Printer's Points. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "2540 HMM (100th Millimeters) converted to Inches = " & $iInch_From_100thMM & @CRLF & _
			"2540 HMM (100th Millimeters) to CM = " & $iCM_From_100thMM & @CRLF & _
			"2540 HMM (100th Millimeters) to MM = " & $iMM_From_100thMM & @CRLF & _
			"2540 HMM (100th Millimeters) to Printer's Points = " & $iPt_From_100thMM)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
