#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#include-once

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Calc Constants for the Libre Office UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various functions in the Libre Office UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; Sleep Divisor $__LOCCONST_SLEEP_DIV
; In applicable functions this is used for adjusting how frequent a sleep occurs in loops.
; For any number above 0 the number of times a loop has completed is divided by $__LOCCONST_SLEEP_DIV. If you find some functions cause momentary freeze ups, a recommended value is 15.
; Set to 0 for no pause in a loop.
Global Const $__LOCCONST_SLEEP_DIV = 0

#Tidy_ILC_Pos=65

; Border Style
Global Const _
		$LOC_BORDERSTYLE_NONE = 0x7FFF, _                       ; No border line.
		$LOC_BORDERSTYLE_SOLID = 0, _                           ; Solid border line.
		$LOC_BORDERSTYLE_DOTTED = 1, _                          ; Dotted border line.
		$LOC_BORDERSTYLE_DASHED = 2, _                          ; Dashed border line.
		$LOC_BORDERSTYLE_DOUBLE = 3, _                          ; Double border line.
		$LOC_BORDERSTYLE_THINTHICK_SMALLGAP = 4, _              ; Double border line with a thin line outside and a thick line inside separated by a small gap.
		$LOC_BORDERSTYLE_THINTHICK_MEDIUMGAP = 5, _             ; Double border line with a thin line outside and a thick line inside separated by a medium gap.
		$LOC_BORDERSTYLE_THINTHICK_LARGEGAP = 6, _              ; Double border line with a thin line outside and a thick line inside separated by a large gap.
		$LOC_BORDERSTYLE_THICKTHIN_SMALLGAP = 7, _              ; Double border line with a thick line outside and a thin line inside separated by a small gap.
		$LOC_BORDERSTYLE_THICKTHIN_MEDIUMGAP = 8, _             ; Double border line with a thick line outside and a thin line inside separated by a medium gap.
		$LOC_BORDERSTYLE_THICKTHIN_LARGEGAP = 9, _              ; Double border line with a thick line outside and a thin line inside separated by a large gap.
		$LOC_BORDERSTYLE_EMBOSSED = 10, _                       ; 3D embossed border line.
		$LOC_BORDERSTYLE_ENGRAVED = 11, _                       ; 3D engraved border line.
		$LOC_BORDERSTYLE_OUTSET = 12, _                         ; Outset border line.
		$LOC_BORDERSTYLE_INSET = 13, _                          ; Inset border line.
		$LOC_BORDERSTYLE_FINE_DASHED = 14, _                    ; Finely dashed border line.
		$LOC_BORDERSTYLE_DOUBLE_THIN = 15, _                    ; Double border line consisting of two fixed thin lines separated by a variable gap.
		$LOC_BORDERSTYLE_DASH_DOT = 16, _                       ; Line consisting of a repetition of one dash and one dot.
		$LOC_BORDERSTYLE_DASH_DOT_DOT = 17                      ; Line consisting of a repetition of one dash and 2 dots.

; Border Width
Global Const _
		$LOC_BORDERWIDTH_HAIRLINE = 2, _                        ; Hairline Border line width.
		$LOC_BORDERWIDTH_VERY_THIN = 18, _                      ; Very Thin Border line width.
		$LOC_BORDERWIDTH_THIN = 26, _                           ; Thin Border line width.
		$LOC_BORDERWIDTH_MEDIUM = 53, _                         ; Medium Border line width.
		$LOC_BORDERWIDTH_THICK = 79, _                          ; Thick Border line width.
		$LOC_BORDERWIDTH_EXTRA_THICK = 159                      ; Extra Thick Border line width.

; Cell Content Horizontal Alignment
Global Const _
	$LOC_CELL_ALIGN_HORI_DEFAULT = 0, _ ; The default alignment is used (left for numbers, right for text).
	$LOC_CELL_ALIGN_HORI_LEFT = 1, _ ; The contents are printed from left to right.
	$LOC_CELL_ALIGN_HORI_CENTER = 2, _ ; The contents are horizontally centered.
	$LOC_CELL_ALIGN_HORI_RIGHT = 3, _ ; The contents are aligned to the right edge of the cell.
	$LOC_CELL_ALIGN_HORI_JUSTIFIED = 4, _ ; The contents are justified to the cell width.
	$LOC_CELL_ALIGN_HORI_FILLED = 5, _ ; The contents are repeated to fill the cell.
	$LOC_CELL_ALIGN_HORI_DISTRIBUTED = 6 ; The contents are evenly aligned across the whole cell. Unlike Justified, it justifies the very last line of text, too.

; Cell Content Vertical Alignment
Global Const _
		$LOC_CELL_ALIGN_VERT_DEFAULT = 0, _ ; The default alignment is used.
		$LOC_CELL_ALIGN_VERT_TOP = 1, _ ; The contents are aligned with the upper edge of the cell.
		$LOC_CELL_ALIGN_VERT_MIDDLE = 2, _ ; The contents are aligned to the vertical middle of the cell.
		$LOC_CELL_ALIGN_VERT_BOTTOM = 3, _ ; The contents are aligned to the lower edge of the cell.
		$LOC_CELL_ALIGN_VERT_JUSTIFIED = 4, _; The contents are justified to the cell height.
		$LOC_CELL_ALIGN_VERT_DISTRIBUTED = 5 ; The same as Justified, unless the text orientation is vertical. Then it behaves similarly to the horizontal Distributed setting, i.e. the very last line is justified, too.

; Cell Delete Mode Constants
Global Const _
		$LOC_CELL_DELETE_MODE_NONE = 0, _ ; No cells are moved -- Nothing happens.
		$LOC_CELL_DELETE_MODE_UP = 1, _ ; The cells below the inserted Cells are moved up.
		$LOC_CELL_DELETE_MODE_LEFT = 2, _ ; The cells to the right of the inserted cells are moved left.
		$LOC_CELL_DELETE_MODE_ROWS = 3, _ ; Entire rows below the inserted cells are moved up.
		$LOC_CELL_DELETE_MODE_COLUMNS = 4 ; Entire columns to the right of the inserted cells are moved left.

; Cell Content Type Flag Constants
Global Const _
		$LOC_CELL_FLAG_VALUE = 1, _ ; Cell Contents that are numeric values but are not formatted as dates or times.
		$LOC_CELL_FLAG_DATE_TIME = 2, _ ; Cell Contents that are numeric values that have a date or time number format.
		$LOC_CELL_FLAG_STRING = 4, _ ; Cell Contents that are Strings.
		$LOC_CELL_FLAG_ANNOTATION = 8, _ ; Cell Annotations
		$LOC_CELL_FLAG_FORMULA = 16, _ ; Cell Contents that are Formulas.
		$LOC_CELL_FLAG_HARD_ATTR = 32, _ ; Cells with explicit formatting, but not the formatting which is applied implicitly through styles.
		$LOC_CELL_FLAG_STYLES = 64, _ ; Cells with Styles applied.
		$LOC_CELL_FLAG_OBJECTS = 128, _ ; Cell Contents that are Drawing Objects.
		$LOC_CELL_FLAG_EDIT_ATTR = 256, _ ; Cells containing formatting within parts of the cell contents.
		$LOC_CELL_FLAG_FORMATTED = 512, _ ; Cells with formatting within the cells or cells with more than one paragraph within the cells.
		$LOC_CELL_FLAG_ALL = 1023 ; All flags listed above.

; Cell Insert Mode Constants
Global Const _
		$LOC_CELL_INSERT_MODE_NONE = 0, _ ; No cells are moved -- Nothing happens.
		$LOC_CELL_INSERT_MODE_DOWN = 1, _ ; The cells below the inserted Cells are moved down.
		$LOC_CELL_INSERT_MODE_RIGHT = 2, _ ; The cells to the right of the inserted cells are moved right.
		$LOC_CELL_INSERT_MODE_ROWS = 3, _ ; Entire rows below the inserted cells are moved down.
		$LOC_CELL_INSERT_MODE_COLUMNS = 4 ; Entire columns to the right of the inserted cells are moved right.

; Cell Content Rotation Reference
Global Const _
		$LOC_CELL_ROTATE_REF_LOWER_CELL_BORDER = 0, _ ; Writes the rotated text from the bottom cell edge outwards.
		$LOC_CELL_ROTATE_REF_UPPER_CELL_BORDER = 1, _ ; Writes the rotated text from the top cell edge outwards.
		$LOC_CELL_ROTATE_REF_INSIDE_CELLS = 3 ; Writes the rotated text only within the cell.

; Cell Type
Global Const _
		$LOC_CELL_TYPE_EMPTY = 0, _                             ; Cell is empty.
		$LOC_CELL_TYPE_VALUE = 1, _                             ; Cell contains a value.
		$LOC_CELL_TYPE_TEXT = 2, _                              ; Cell contains text.
		$LOC_CELL_TYPE_FORMULA = 3                              ; Cell contains a formula.

; Color in Long Color Format
Global Const _
		$LOC_COLOR_OFF = -1, _                                  ; Turn Color off, or to automatic mode.
		$LOC_COLOR_BLACK = 0, _                                 ; Black color.
		$LOC_COLOR_WHITE = 16777215, _                          ; White color.
		$LOC_COLOR_LGRAY = 11711154, _                          ; Light Gray color.
		$LOC_COLOR_GRAY = 8421504, _                            ; Gray color.
		$LOC_COLOR_DKGRAY = 3355443, _                          ; Dark Gray color.
		$LOC_COLOR_YELLOW = 16776960, _                         ; Yellow color.
		$LOC_COLOR_GOLD = 16760576, _                           ; Gold color.
		$LOC_COLOR_ORANGE = 16744448, _                         ; Orange color.
		$LOC_COLOR_BRICK = 16728064, _                          ; Brick color.
		$LOC_COLOR_RED = 16711680, _                            ; Red color.
		$LOC_COLOR_MAGENTA = 12517441, _                        ; Magenta color.
		$LOC_COLOR_PURPLE = 8388736, _                          ; Purple color.
		$LOC_COLOR_INDIGO = 5582989, _                          ; Indigo color.
		$LOC_COLOR_BLUE = 2777241, _                            ; Blue color.
		$LOC_COLOR_TEAL = 1410150, _                            ; Teal color.
		$LOC_COLOR_GREEN = 43315, _                             ; Green color.
		$LOC_COLOR_LIME = 8508442, _                            ; Lime color.
		$LOC_COLOR_BROWN = 9127187                              ; Brown color.

; Formula Result Type Constants
Global Const _
		$LOC_FORMULA_RESULT_TYPE_VALUE = 1, _ ; The formula's result is a number.
		$LOC_FORMULA_RESULT_TYPE_STRING = 2, _ ; The formula's result is a string.
		$LOC_FORMULA_RESULT_TYPE_ERROR = 4, _ ; The formula has an error of some form.
		$LOC_FORMULA_RESULT_TYPE_ALL = 7 ; All of the above types.

; Path Convert Constants.
Global Const _
		$LOC_PATHCONV_AUTO_RETURN = 0, _                        ; Automatically returns the opposite of the input path, determined by StringInStr search for either "File:///"(L.O.Office URL) or "[A-Z]:\" (Windows File Path).
		$LOC_PATHCONV_OFFICE_RETURN = 1, _                      ; Returns L.O. Office URL, even if the input is already in that format.
		$LOC_PATHCONV_PCPATH_RETURN = 2                         ; Returns Windows File Path, even if the input is already in that format.

; Posture/Italic
Global Const _
		$LOC_POSTURE_NONE = 0, _                                ; Specifies a font without slant.
		$LOC_POSTURE_OBLIQUE = 1, _                             ; Specifies an oblique font (slant not designed into the font).
		$LOC_POSTURE_ITALIC = 2, _                              ; Specifies an italic font (slant designed into the font).
		$LOC_POSTURE_DontKnow = 3, _                            ; Specifies a font with an unknown slant. For Read Only.
		$LOC_POSTURE_REV_OBLIQUE = 4, _                         ; Specifies a reverse oblique font (slant not designed into the font).
		$LOC_POSTURE_REV_ITALIC = 5                             ; Specifies a reverse italic font (slant designed into the font).

; Relief
Global Const _
		$LOC_RELIEF_NONE = 0, _                                 ; No relief is applied.
		$LOC_RELIEF_EMBOSSED = 1, _                             ; The font relief is embossed.
		$LOC_RELIEF_ENGRAVED = 2                                ; The font relief is engraved.

; Shadow Location
Global Const _
		$LOC_SHADOW_NONE = 0, _                                 ; No shadow is applied.
		$LOC_SHADOW_TOP_LEFT = 1, _                             ; Shadow is located along the upper and left sides.
		$LOC_SHADOW_TOP_RIGHT = 2, _                            ; Shadow is located along the upper and right sides.
		$LOC_SHADOW_BOTTOM_LEFT = 3, _                          ; Shadow is located along the lower and left sides.
		$LOC_SHADOW_BOTTOM_RIGHT = 4                            ; Shadow is located along the lower and right sides.

; Strikeout
Global Const _
		$LOC_STRIKEOUT_NONE = 0, _                              ; No strike out.
		$LOC_STRIKEOUT_SINGLE = 1, _                            ; Strike out the characters with a single line.
		$LOC_STRIKEOUT_DOUBLE = 2, _                            ; Strike out the characters with a double line.
		$LOC_STRIKEOUT_DONT_KNOW = 3, _                         ; The strikeout mode is not specified. For Read Only.
		$LOC_STRIKEOUT_BOLD = 4, _                              ; Strike out the characters with a bold line.
		$LOC_STRIKEOUT_SLASH = 5, _                             ; Strike out the characters with slashes.
		$LOC_STRIKEOUT_X = 6                                    ; Strike out the characters with X's.

; Text Direction
Global Const _
		$LOC_TXT_DIR_LR = 0, _                               ; Text within lines is written left-to-right. Typically, this is the writing mode for normal "alphabetic" text.
		$LOC_TXT_DIR_RL = 1, _                               ; Text within a line are written right-to-left. Typically, this writing mode is used in Arabic and Hebrew text.
		$LOC_TXT_DIR_CONTEXT = 4                            ; Obtain actual writing mode from the context of the object.

; Underline/Overline
Global Const _
		$LOC_UNDERLINE_NONE = 0, _                              ; No Underline or Overline style.
		$LOC_UNDERLINE_SINGLE = 1, _                            ; Single line Underline/Overline style.
		$LOC_UNDERLINE_DOUBLE = 2, _                            ; Double line Underline/Overline style.
		$LOC_UNDERLINE_DOTTED = 3, _                            ; Dotted line Underline/Overline style.
		$LOC_UNDERLINE_DONT_KNOW = 4, _                         ; Unknown Underline/Overline style, for read only.
		$LOC_UNDERLINE_DASH = 5, _                              ; Dashed line Underline/Overline style.
		$LOC_UNDERLINE_LONG_DASH = 6, _                         ; Long Dashed line Underline/Overline style.
		$LOC_UNDERLINE_DASH_DOT = 7, _                          ; Dash Dot line Underline/Overline style.
		$LOC_UNDERLINE_DASH_DOT_DOT = 8, _                      ; Dash Dot Dot line Underline/Overline style.
		$LOC_UNDERLINE_SML_WAVE = 9, _                          ; Small Wave line Underline/Overline style.
		$LOC_UNDERLINE_WAVE = 10, _                             ; Wave line Underline/Overline style.
		$LOC_UNDERLINE_DBL_WAVE = 11, _                         ; Double Wave line Underline/Overline style.
		$LOC_UNDERLINE_BOLD = 12, _                             ; Bold line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_DOTTED = 13, _                      ; Bold Dotted line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_DASH = 14, _                        ; Bold Dashed line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_LONG_DASH = 15, _                   ; Bold Long Dash line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_DASH_DOT = 16, _                    ; Bold Dash Dot line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_DASH_DOT_DOT = 17, _                ; Bold Dash Dot Dot line Underline/Overline style.
		$LOC_UNDERLINE_BOLD_WAVE = 18                           ; Bold Wave line Underline/Overline style.

; Weight/Bold
Global Const _
		$LOC_WEIGHT_DONT_KNOW = 0, _                            ; The font weight is not specified/unknown. For Read Only.
		$LOC_WEIGHT_THIN = 50, _                                ; A 50% (Thin) font weight.
		$LOC_WEIGHT_ULTRA_LIGHT = 60, _                         ; A 60% (Ultra Light) font weight.
		$LOC_WEIGHT_LIGHT = 75, _                               ; A 75% (Light) font weight.
		$LOC_WEIGHT_SEMI_LIGHT = 90, _                          ; A 90% (Semi-Light) font weight.
		$LOC_WEIGHT_NORMAL = 100, _                             ; A 100% (Normal) font weight.
		$LOC_WEIGHT_SEMI_BOLD = 110, _                          ; A 110% (Semi-Bold) font weight.
		$LOC_WEIGHT_BOLD = 150, _                               ; A 150% (Bold) font weight.
		$LOC_WEIGHT_ULTRA_BOLD = 175, _                         ; A 175% (Ultra-Bold) font weight.
		$LOC_WEIGHT_BLACK = 200                                 ; A 200% (Black) font weight.

; Zoom Type Constants
Global Const _
		$LOC_ZOOMTYPE_OPTIMAL = 0, _                            ; The page content width (excluding margins) at the current selection is fit into the view.
		$LOC_ZOOMTYPE_PAGE_WIDTH = 1, _                         ; The page width at the current selection is fit into the view.
		$LOC_ZOOMTYPE_ENTIRE_PAGE = 2, _                        ; A complete page of the document is fit into the view.
		$LOC_ZOOMTYPE_BY_VALUE = 3, _                           ; The Zoom property is relative, and set using Zoom Value.
		$LOC_ZOOMTYPE_PAGE_WIDTH_EXACT = 4                      ; The Page width at the current selection is fit into the view with the view ends exactly at the end of the page.
