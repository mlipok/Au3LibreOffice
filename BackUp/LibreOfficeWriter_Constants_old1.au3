#include-once

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter) Constants for the Libre Office Writer UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various functions in the Libre Office Writer UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; Sleep Divisor $__LOWCONST_SLEEP_DIV
; Lower this number for more frequent sleeps in applicable functions, raise it for less.
; Set to 0 for no pause in a loop.
Global Const $__LOWCONST_SLEEP_DIV = 15

#Tidy_ILC_Pos=60
; Error Codes
Global Enum _
		$__LOW_STATUS_SUCCESS = 0, _                       ; 0
		$__LOW_STATUS_INPUT_ERROR, _                       ; 1
		$__LOW_STATUS_INIT_ERROR, _                        ; 2
		$__LOW_STATUS_PROCESSING_ERROR, _                  ; 3
		$__LOW_STATUS_PROP_SETTING_ERROR, _                ; 4
		$__LOW_STATUS_DOC_ERROR, _                         ; 5
		$__LOW_STATUS_PRINTER_RELATED_ERROR, _             ; 6
		$__LOW_STATUS_VER_ERROR                            ; 7

; Conversion Constants.
Global Enum _
		$__LOWCONST_CONVERT_TWIPS_CM, _                    ; 0
		$__LOWCONST_CONVERT_TWIPS_INCH, _                  ; 1
		$__LOWCONST_CONVERT_TWIPS_UM, _                    ; 2
		$__LOWCONST_CONVERT_MM_UM, _                       ; 3
		$__LOWCONST_CONVERT_UM_MM, _                       ; 4
		$__LOWCONST_CONVERT_CM_UM, _                       ; 5
		$__LOWCONST_CONVERT_UM_CM, _                       ; 6
		$__LOWCONST_CONVERT_INCH_UM, _                     ; 7
		$__LOWCONST_CONVERT_UM_INCH, _                     ; 8
		$__LOWCONST_CONVERT_PT_UM, _                       ; 9
		$__LOWCONST_CONVERT_UM_PT                          ; 10

; Fill Style Type Constants
Global Enum _
		$__LOWCONST_FILL_STYLE_OFF, _                      ; 0
		$__LOWCONST_FILL_STYLE_SOLID, _                    ; 1
		$__LOWCONST_FILL_STYLE_GRADIENT, _                 ; 2
		$__LOWCONST_FILL_STYLE_HATCH, _                    ; 3
		$__LOWCONST_FILL_STYLE_BITMAP                      ; 4

; Cursor Data Related Constants
Global Const _
		$LOW_CURDATA_BODY_TEXT = 1, _                      ;Cursor is currently in the Body Text.
		$LOW_CURDATA_FRAME = 2, _                          ;Cursor is currently in a Text Frame.
		$LOW_CURDATA_CELL = 3, _                           ;Cursor is currently in a Text Table Cell.
		$LOW_CURDATA_FOOTNOTE = 4, _                       ;Cursor is currently in a Footnote.
		$LOW_CURDATA_ENDNOTE = 5, _                        ;Cursor is currently in a Endnote.
		$LOW_CURDATA_HEADER_FOOTER = 6                     ;Cursor is currently in a Header or Footer.

; Cursor Type Related Constants
Global Const _
		$LOW_CURTYPE_TEXT_CURSOR = 1, _                    ;Cursor is a TextCursor type.
		$LOW_CURTYPE_TABLE_CURSOR = 2, _                   ;Cursor is a TableCursor type.
		$LOW_CURTYPE_VIEW_CURSOR = 3, _                    ;Cursor is a ViewCursor type.
		$LOW_CURTYPE_PARAGRAPH = 4, _                      ;Object is a Paragraph Object.
		$LOW_CURTYPE_TEXT_PORTION = 5                      ;Object is a Paragraph Text Portion Object.

; Path Convert Constants.
Global Const _
		$LOW_PATHCONV_AUTO_RETURN = 0, _                   ; Automatically returns the opposite of the input.
		$LOW_PATHCONV_OFFICE_RETURN = 1, _                 ; Returns L.O. Office URL.
		$LOW_PATHCONV_PCPATH_RETURN = 2                    ; Returns Windows File Path.

; Printer Duplex Constants.
Global Const _
		$LOW_DUPLEX_UNKNOWN = 0, _                         ; Duplex mode setting is unknown.
		$LOW_DUPLEX_OFF = 1, _                             ; Duplex mode is off.
		$LOW_DUPLEX_LONG = 2, _                            ; Duplex mode is on, flip on Long edge.
		$LOW_DUPLEX_SHORT = 3                              ; Duplex mode is on, flip on Short edge.

; Printer Paper Orientation Constants.
Global Const _
		$LOW_PAPER_PORTRAIT = 0, _                         ; Portrait Paper Orientation.
		$LOW_PAPER_LANDSCAPE = 1                           ; Landscape Paper Orientation.

; Paper Size Constants.
Global Const _
		$LOW_PAPER_A3 = 0, _                               ; A3 Paper size.
		$LOW_PAPER_A4 = 1, _                               ; A4 Paper size.
		$LOW_PAPER_A5 = 2, _                               ; A5 Paper size.
		$LOW_PAPER_B4 = 3, _                               ; B4 Paper size.
		$LOW_PAPER_B5 = 4, _                               ; B5 Paper size.
		$LOW_PAPER_LETTER = 5, _                           ; Letter Paper size.
		$LOW_PAPER_LEGAL = 6, _                            ; Legal Paper size.
		$LOW_PAPER_TABLOID = 7, _                          ; Tabloid Paper size.
		$LOW_PAPER_USER_DEFINED = 8                        ; Paper size is User-Defined.

; LO Print Comments Constants.
Global Const _
		$LOW_PRINT_NOTES_NONE = 0, _                       ; Document contents are printed, without printing any Comments.
		$LOW_PRINT_NOTES_ONLY = 1, _                       ; Only Comments are printed, and NONE of the Document content.
		$LOW_PRINT_NOTES_END = 2, _                        ; Document content is printed with comments appended to a blank page at the end of the document.
		$LOW_PRINT_NOTES_NEXT_PAGE = 3                     ; Document content is printed and comments are appended to a blank page after the commented page.

; LO ViewCursor Movement Constants.
Global Enum _
		$LOW_VIEWCUR_GO_DOWN, _                            ; Move the cursor Down.
		$LOW_VIEWCUR_GO_UP, _                              ; Move the cursor Up.
		$LOW_VIEWCUR_GO_LEFT, _                            ; Move the cursor left.
		$LOW_VIEWCUR_GO_RIGHT, _                           ; Move the cursor right.
		$LOW_VIEWCUR_GOTO_END_OF_LINE, _                   ;     Move the cursor to the end of the current line.
		$LOW_VIEWCUR_GOTO_START_OF_LINE, _                 ; Move the cursor to the start of the current line.
		$LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, _                 ; Move the cursor to the first page.
		$LOW_VIEWCUR_JUMP_TO_LAST_PAGE, _                  ; Move the cursor to the Last page.
		$LOW_VIEWCUR_JUMP_TO_PAGE, _                       ; Jump to a specified page.
		$LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, _                  ; Move the cursor to the Next page.
		$LOW_VIEWCUR_JUMP_TO_PREV_PAGE, _                  ; Move the cursor to the previous page.
		$LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, _                ; Move the cursor to the end of the current page.
		$LOW_VIEWCUR_JUMP_TO_START_OF_PAGE, _              ; Move the cursor to the start of the current page.
		$LOW_VIEWCUR_SCREEN_DOWN, _                        ; Scroll the view forward by one visible page.
		$LOW_VIEWCUR_SCREEN_UP, _                          ; Scroll the view back by one visible page.
		$LOW_VIEWCUR_GOTO_START, _                         ; Move the cursor to the start of the document or Table.
		$LOW_VIEWCUR_GOTO_END                              ; Move the cursor to the end of the document or Table.

; LO TextCursor Movement Constants.
Global Enum _
		$LOW_TEXTCUR_COLLAPSE_TO_START, _                  ; Collapses the current selection the start of the selection.
		$LOW_TEXTCUR_COLLAPSE_TO_END, _                    ; Collapses the current selection the end of the selection.
		$LOW_TEXTCUR_GO_LEFT, _                            ; Move the cursor left.
		$LOW_TEXTCUR_GO_RIGHT, _                           ; Move the cursor right.
		$LOW_TEXTCUR_GOTO_START, _                         ; Move the cursor to the start of the text.
		$LOW_TEXTCUR_GOTO_END, _                           ; Move the cursor to the end of the text.
		$LOW_TEXTCUR_GOTO_NEXT_WORD, _                     ; Move to the start of the next word.
		$LOW_TEXTCUR_GOTO_PREV_WORD, _                     ; Move to the end of the previous word.
		$LOW_TEXTCUR_GOTO_END_OF_WORD, _                   ; Move to the end of the current word.
		$LOW_TEXTCUR_GOTO_START_OF_WORD, _                 ; Move to the start of the current word.
		$LOW_TEXTCUR_GOTO_NEXT_SENTENCE, _                 ; Move to the start of the next sentence.
		$LOW_TEXTCUR_GOTO_PREV_SENTENCE, _                 ; Move to the end of the previous sentence.
		$LOW_TEXTCUR_GOTO_END_OF_SENTENCE, _               ; Move to the end of the current sentence.
		$LOW_TEXTCUR_GOTO_START_OF_SENTENCE, _             ; Move to the start of the current sentence.
		$LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH, _                ; Move to the start of the next paragraph.
		$LOW_TEXTCUR_GOTO_PREV_PARAGRAPH, _                ; Move to the End of the previous paragraph.
		$LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH, _              ; Move to the end of the current paragraph.
		$LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH               ; Move to the start of the current paragraph.

; LO TableCursor Movement Constants.
Global Enum _
		$LOW_TABLECUR_GO_LEFT, _                           ; Move the cursor left.
		$LOW_TABLECUR_GO_RIGHT, _                          ; Move the cursor right.
		$LOW_TABLECUR_GO_UP, _                             ; Move the cursor up.
		$LOW_TABLECUR_GO_DOWN, _                           ; Move the cursor down.
		$LOW_TABLECUR_GOTO_START, _                        ; Move the cursor to the first cell.
		$LOW_TABLECUR_GOTO_END                             ; Move the cursor to the last cell.

; Break Type
Global Const _
		$LOW_BREAK_NONE = 0, _                             ; No column or page break is applied.
		$LOW_BREAK_COLUMN_BEFORE = 1, _                    ; A column break is applied before the current Paragraph.
		$LOW_BREAK_COLUMN_AFTER = 2, _                     ; A column break is applied after the current Paragraph.
		$LOW_BREAK_COLUMN_BOTH = 3, _                      ; A column break is applied before and after the current Paragraph.
		$LOW_BREAK_PAGE_BEFORE = 4, _                      ; A page break is applied before the current Paragraph.
		$LOW_BREAK_PAGE_AFTER = 5, _                       ; A page break is applied after the current Paragraph.
		$LOW_BREAK_PAGE_BOTH = 6                           ; A page break is applied before and after the current Paragraph.

; Horizontal Orientation
Global Const _
		$LOW_ORIENT_HORI_NONE = 0, _                       ; No hard alignment is applied. Equal to "From Left" in L.O. U.I.
		$LOW_ORIENT_HORI_RIGHT = 1, _                      ; The object is aligned at the right side.
		$LOW_ORIENT_HORI_CENTER = 2, _                     ; The object is aligned at the middle.
		$LOW_ORIENT_HORI_LEFT = 3, _                       ; The object is aligned at the left side.
		$LOW_ORIENT_HORI_FULL = 6, _                       ; The table uses the full space (for text tables only).
		$LOW_ORIENT_HORI_LEFT_AND_WIDTH = 7                ;  The left offset and the width of the table are defined.

; Color in Long Color Format
Global Const _
		$LOW_COLOR_OFF = -1, _                             ; Turn Color off, or to automatic mode.
		$LOW_COLOR_BLACK = 0, _                            ; Black color.
		$LOW_COLOR_WHITE = 16777215, _                     ; White color.
		$LOW_COLOR_LGRAY = 11711154, _                     ; Light Gray color.
		$LOW_COLOR_GRAY = 8421504, _                       ; Gray color.
		$LOW_COLOR_DKGRAY = 3355443, _                     ; Dark Gray color.
		$LOW_COLOR_YELLOW = 16776960, _                    ; Yellow color.
		$LOW_COLOR_GOLD = 16760576, _                      ; Gold color.
		$LOW_COLOR_ORANGE = 16744448, _                    ; Orange color.
		$LOW_COLOR_BRICK = 16728064, _                     ; Brick color.
		$LOW_COLOR_RED = 16711680, _                       ; Red color.
		$LOW_COLOR_MAGENTA = 12517441, _                   ; Magenta color.
		$LOW_COLOR_PURPLE = 8388736, _                     ; Purple color.
		$LOW_COLOR_INDIGO = 5582989, _                     ; Indigo color.
		$LOW_COLOR_BLUE = 2777241, _                       ; Blue color.
		$LOW_COLOR_TEAL = 1410150, _                       ; Teal color.
		$LOW_COLOR_GREEN = 43315, _                        ; Green color.
		$LOW_COLOR_LIME = 8508442, _                       ; Lime color.
		$LOW_COLOR_BROWN = 9127187                         ; Brown color.

; Border Style
Global Const _
		$LOW_BORDERSTYLE_NONE = 0x7FFF, _                  ; No border line.
		$LOW_BORDERSTYLE_SOLID = 0, _                      ; Solid border line.
		$LOW_BORDERSTYLE_DOTTED = 1, _                     ; Dotted border line.
		$LOW_BORDERSTYLE_DASHED = 2, _                     ; Dashed border line.
		$LOW_BORDERSTYLE_DOUBLE = 3, _                     ; Double border line.
		$LOW_BORDERSTYLE_THINTHICK_SMALLGAP = 4, _         ; Double border line with a thin line outside and a thick line inside separated by a small gap.
		$LOW_BORDERSTYLE_THINTHICK_MEDIUMGAP = 5, _        ; Double border line with a thin line outside and a thick line inside separated by a medium gap.
		$LOW_BORDERSTYLE_THINTHICK_LARGEGAP = 6, _         ; Double border line with a thin line outside and a thick line inside separated by a large gap.
		$LOW_BORDERSTYLE_THICKTHIN_SMALLGAP = 7, _         ; Double border line with a thick line outside and a thin line inside separated by a small gap.
		$LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP = 8, _        ; Double border line with a thick line outside and a thin line inside separated by a medium gap.
		$LOW_BORDERSTYLE_THICKTHIN_LARGEGAP = 9, _         ; Double border line with a thick line outside and a thin line inside separated by a large gap.
		$LOW_BORDERSTYLE_EMBOSSED = 10, _                  ; 3D embossed border line.
		$LOW_BORDERSTYLE_ENGRAVED = 11, _                  ; 3D engraved border line.
		$LOW_BORDERSTYLE_OUTSET = 12, _                    ; Outset border line.
		$LOW_BORDERSTYLE_INSET = 13, _                     ; Inset border line.
		$LOW_BORDERSTYLE_FINE_DASHED = 14, _               ; Finely dashed border line.
		$LOW_BORDERSTYLE_DOUBLE_THIN = 15, _               ; Double border line consisting of two fixed thin lines separated by a variable gap.
		$LOW_BORDERSTYLE_DASH_DOT = 16, _                  ; Line consisting of a repetition of one dash and one dot.
		$LOW_BORDERSTYLE_DASH_DOT_DOT = 17                 ; Line consisting of a repetition of one dash and 2 dots.

; Border Width
Global Const _
		$LOW_BORDERWIDTH_HAIRLINE = 2, _                   ; Hairline Border line width.
		$LOW_BORDERWIDTH_VERY_THIN = 18, _                 ; Very Thin Border line width.
		$LOW_BORDERWIDTH_THIN = 26, _                      ; Thin Border line width.
		$LOW_BORDERWIDTH_MEDIUM = 53, _                    ; Medium Border line width.
		$LOW_BORDERWIDTH_THICK = 79, _                     ; Thick Border line width.
		$LOW_BORDERWIDTH_EXTRA_THICK = 159                 ; Extra Thick Border line width.

; Vertical Orientation
Global Const _
		$LOW_ORIENT_VERT_NONE = 0, _                       ; No hard alignment. The same as "From Top"/From Bottom" in L.O. U.I.
		$LOW_ORIENT_VERT_TOP = 1, _                        ; Aligned at the top.
		$LOW_ORIENT_VERT_CENTER = 2, _                     ; Aligned at the center.
		$LOW_ORIENT_VERT_BOTTOM = 3, _                     ; Aligned at the bottom.
		$LOW_ORIENT_VERT_CHAR_TOP = 4, _                   ; Aligned at the top of a character.
		$LOW_ORIENT_VERT_CHAR_CENTER = 5, _                ; Aligned at the center of a character.
		$LOW_ORIENT_VERT_CHAR_BOTTOM = 6, _                ; Aligned at the bottom of a character.
		$LOW_ORIENT_VERT_LINE_TOP = 7, _                   ; Aligned at the top of the line.
		$LOW_ORIENT_VERT_LINE_CENTER = 8, _                ; Aligned at the center of the line.
		$LOW_ORIENT_VERT_LINE_BOTTOM = 9                   ; Aligned at the bottom of the line.

; Tab Alignment
Global Const _
		$LOW_TAB_ALIGN_LEFT = 0, _                         ; Aligns the left edge of the text to the tab stop and extends the text to the right.
		$LOW_TAB_ALIGN_CENTER = 1, _                       ; Aligns the center of the text to the tab stop.
		$LOW_TAB_ALIGN_RIGHT = 2, _                        ; Aligns the right edge of the text to the tab stop and extends the text to the left of the tab stop.
		$LOW_TAB_ALIGN_DECIMAL = 3, _                      ; Aligns the decimal separator of a number to the center of the tab stop and text to the left of the tab.
		$LOW_TAB_ALIGN_DEFAULT = 4                         ; This setting is the default, setting when no TabStops are present.

Global Const _ ; Underline/Overline
		$LOW_UNDERLINE_NONE = 0, _ ; No Underline or Overline style.
		$LOW_UNDERLINE_SINGLE = 1, _ ; Single line Underline/Overline style.
		$LOW_UNDERLINE_DOUBLE = 2, _ ; Double line Underline/Overline style.
		$LOW_UNDERLINE_DOTTED = 3, _ ; Dotted line Underline/Overline style.
		$LOW_UNDERLINE_DONT_KNOW = 4, _ ; Unknown Underline/Overline style, for read only.
		$LOW_UNDERLINE_DASH = 5, _ ; Dashed line Underline/Overline style.
		$LOW_UNDERLINE_LONG_DASH = 6, _ ; Long Dashed line Underline/Overline style.
		$LOW_UNDERLINE_DASH_DOT = 7, _ ; Dash Dot line Underline/Overline style.
		$LOW_UNDERLINE_DASH_DOT_DOT = 8, _ ; Dash Dot Dot line Underline/Overline style.
		$LOW_UNDERLINE_SML_WAVE = 9, _ ; Small Wave line Underline/Overline style.
		$LOW_UNDERLINE_WAVE = 10, _ ; Wave line Underline/Overline style.
		$LOW_UNDERLINE_DBL_WAVE = 11, _ ; Double Wave line Underline/Overline style.
		$LOW_UNDERLINE_BOLD = 12, _ ; Bold line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_DOTTED = 13, _ ; Bold Dotted line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_DASH = 14, _ ; Bold Dashed line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_LONG_DASH = 15, _ ; Bold Long Dash line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_DASH_DOT = 16, _ ; Bold Dash Dot line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_DASH_DOT_DOT = 17, _ ; Bold Dash Dot Dot line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_WAVE = 18 ; Bold Wave line Underline/Overline style.

Global Const _ ; Strikeout
		$LOW_STRIKEOUT_NONE = 0, _ ; No strike out.
		$LOW_STRIKEOUT_SINGLE = 1, _ ; Strike out the characters with a single line.
		$LOW_STRIKEOUT_DOUBLE = 2, _ ; Strike out the characters with a double line.
		$LOW_STRIKEOUT_DONT_KNOW = 3, _ ; The strikeout mode is not specified. For Read Only.
		$LOW_STRIKEOUT_BOLD = 4, _ ; Strike out the characters with a bold line.
		$LOW_STRIKEOUT_SLASH = 5, _ ; Strike out the characters with slashes.
		$LOW_STRIKEOUT_X = 6 ; Strike out the characters with X's.

Global Const _ ; Relief
		$LOW_RELIEF_NONE = 0, _ ; No relief is applied.
		$LOW_RELIEF_EMBOSSED = 1, _ ; The font relief is embossed.
		$LOW_RELIEF_ENGRAVED = 2 ; The font relief is engraved.

Global Const _ ; Case
		$LOW_CASEMAP_NONE = 0, _ ; The case of the characters is unchanged.
		$LOW_CASEMAP_UPPER = 1, _ ; All characters are put in upper case.
		$LOW_CASEMAP_LOWER = 2, _ ; All characters are put in lower case.
		$LOW_CASEMAP_TITLE = 3, _ ; The first character of each word is put in upper case.
		$LOW_CASEMAP_SM_CAPS = 4 ; All characters are put in upper case, but with a smaller font height.

Global Const _ ; Shadow
		$LOW_SHADOW_NONE = 0, _ ; No shadow is applied.
		$LOW_SHADOW_TOP_LEFT = 1, _ ; Shadow is located along the upper and left sides.
		$LOW_SHADOW_TOP_RIGHT = 2, _ ; Shadow is located along the upper and right sides.
		$LOW_SHADOW_BOTTOM_LEFT = 3, _ ; Shadow is located along the lower and left sides.
		$LOW_SHADOW_BOTTOM_RIGHT = 4 ; Shadow is located along the lower and right sides.

Global Const _ ; Posture/Italic
		$LOW_POSTURE_NONE = 0, _ ; Specifies a font without slant.
		$LOW_POSTURE_OBLIQUE = 1, _ ; Specifies an oblique font (slant not designed into the font).
		$LOW_POSTURE_ITALIC = 2, _ ; Specifies an italic font (slant designed into the font).
		$LOW_POSTURE_DontKnow = 3, _ ; Specifies a font with an unknown slant. For Read Only.
		$LOW_POSTURE_REV_OBLIQUE = 4, _ ; Specifies a reverse oblique font (slant not designed into the font).
		$LOW_POSTURE_REV_ITALIC = 5 ; Specifies a reverse italic font (slant designed into the font).

Global Const _ ; Weight/Bold
		$LOW_WEIGHT_DONT_KNOW = 0, _ ; The font weight is not specified/unknown. For Read Only.
		$LOW_WEIGHT_THIN = 50, _ ; A 50% (Thin) font weight.
		$LOW_WEIGHT_ULTRA_LIGHT = 60, _ ; A 60% (Ultra Light) font weight.
		$LOW_WEIGHT_LIGHT = 75, _ ; A 75% (Light) font weight.
		$LOW_WEIGHT_SEMI_LIGHT = 90, _ ; A 90% (Semi-Light) font weight.
		$LOW_WEIGHT_NORMAL = 100, _ ; A 100% (Normal) font weight.
		$LOW_WEIGHT_SEMI_BOLD = 110, _ ; A 110% (Semi-Bold) font weight.
		$LOW_WEIGHT_BOLD = 150, _ ; A 150% (Bold) font weight.
		$LOW_WEIGHT_ULTRA_BOLD = 175, _ ; A 175% (Ultra-Bold) font weight.
		$LOW_WEIGHT_BLACK = 200 ; A 200% (Black) font weight.

Global Const _ ; Outline
		$LOW_OUTLINE_BODY = 0, _ ; The paragraph belongs to the body text.
		$LOW_OUTLINE_LEVEL_1 = 1, _ ; The paragraph belongs to the level 1 outline level.
		$LOW_OUTLINE_LEVEL_2 = 2, _ ; The paragraph belongs to the level 2 outline level.
		$LOW_OUTLINE_LEVEL_3 = 3, _ ; The paragraph belongs to the level 3 outline level.
		$LOW_OUTLINE_LEVEL_4 = 4, _ ; The paragraph belongs to the level 4 outline level.
		$LOW_OUTLINE_LEVEL_5 = 5, _ ; The paragraph belongs to the level 5 outline level.
		$LOW_OUTLINE_LEVEL_6 = 6, _ ; The paragraph belongs to the level 6 outline level.
		$LOW_OUTLINE_LEVEL_7 = 7, _ ; The paragraph belongs to the level 7 outline level.
		$LOW_OUTLINE_LEVEL_8 = 8, _ ; The paragraph belongs to the level 8 outline level.
		$LOW_OUTLINE_LEVEL_9 = 9, _ ; The paragraph belongs to the level 9 outline level.
		$LOW_OUTLINE_LEVEL_10 = 10 ; The paragraph belongs to the level 10 outline level.

Global Const _ ; Line Spacing
		$LOW_LINE_SPC_MODE_PROP = 0, _ ; Specifies the height value as a proportional value.
		$LOW_LINE_SPC_MODE_MIN = 1, _ ; Specifies the height as the minimum line height.
		$LOW_LINE_SPC_MODE_LEADING = 2, _ ; Specifies the height value as the distance to the previous line.
		$LOW_LINE_SPC_MODE_FIX = 3 ; Specifies the height value as a fixed line height.

Global Const _ ; Paragraph Horizontal Align
		$LOW_PAR_ALIGN_HOR_LEFT = 0, _ ; The Paragraph is left-aligned between the borders.
		$LOW_PAR_ALIGN_HOR_RIGHT = 1, _ ; The Paragraph is right-aligned between the borders.
		$LOW_PAR_ALIGN_HOR_JUSTIFIED = 2, _ ; The Paragraph is adjusted / stretched to both borders.
		$LOW_PAR_ALIGN_HOR_CENTER = 3, _ ; The Paragraph is centered between the left and right borders.
		$LOW_PAR_ALIGN_HOR_STRETCH = 4 ;HoriAlign 4 does nothing??

Global Const _ ; Paragraph Vertical Align
		$LOW_PAR_ALIGN_VERT_AUTO = 0, _ ; Automatic vertical alignment mode.
		$LOW_PAR_ALIGN_VERT_BASELINE = 1, _ ; The text is aligned to the baseline.
		$LOW_PAR_ALIGN_VERT_TOP = 2, _ ; The text is aligned to the top.
		$LOW_PAR_ALIGN_VERT_CENTER = 3, _ ; The text is aligned to the center.
		$LOW_PAR_ALIGN_VERT_BOTTOM = 4 ; The text is aligned to bottom.

Global Const _ ; Paragraph Last Line Alignment
		$LOW_PAR_LAST_LINE_START = 0, _ ; The Paragraph is aligned either to the Left border or the right, depending on the current text direction.
		$LOW_PAR_LAST_LINE_JUSTIFIED = 2, _ ; The Paragraph is adjusted to both borders / stretched.
		$LOW_PAR_LAST_LINE_CENTER = 3 ; The Paragraph is centered between the left and right borders.

Global Const _ ; Text Direction
		$LOW_TXT_DIR_LR_TB = 0, _ ; Text within lines is written left-to-right. Lines and blocks are placed top-to-bottom.
		$LOW_TXT_DIR_RL_TB = 1, _ ; Text within a line are written right-to-left. Lines and blocks are placed top-to-bottom.
		$LOW_TXT_DIR_TB_RL = 2, _ ; Text within a line is written top-to-bottom. Lines and blocks are placed right-to-left.
		$LOW_TXT_DIR_TB_LR = 3, _ ; Text within a line is written top-to-bottom. Lines and blocks are placed left-to-right.
		$LOW_TXT_DIR_CONTEXT = 4, _ ; Obtain actual writing mode from the context of the object.
		$LOW_TXT_DIR_BT_LR = 5 ; text within a line is written bottom-to-top. Lines and blocks are placed left-to-right. (LibreOffice 6.3).

Global Const _ ; Control Character
		$LOW_CON_CHAR_PAR_BREAK = 0, _ ; A new paragraph.
		$LOW_CON_CHAR_LINE_BREAK = 1, _ ; A new line in a paragraph.
		$LOW_CON_CHAR_HARD_HYPHEN = 2, _ ; A dash but prevents this position from being hyphenated.
		$LOW_CON_CHAR_SOFT_HYPHEN = 3, _ ; Defines a preferred hyphenation point if the word must be split at the end of a line.
		$LOW_CON_CHAR_HARD_SPACE = 4, _ ; Insert a space that prevents two words from splitting at a line break.
		$LOW_CON_CHAR_APPEND_PAR = 5 ; Appends a new paragraph.

Global Const _ ; Cell Type
		$LOW_CELL_TYPE_EMPTY = 0, _ ; Cell is empty.
		$LOW_CELL_TYPE_VALUE = 1, _ ; Cell contains a value.
		$LOW_CELL_TYPE_TEXT = 2, _ ; Cell contains text.
		$LOW_CELL_TYPE_FORMULA = 3 ; Cell contains a formula.

Global Const _ ; Paper Width in uM
		$LOW_PAPER_WIDTH_A6 = 10490, _ ; A6 paper width in Micrometers.
		$LOW_PAPER_WIDTH_A5 = 14808, _ ; A5 paper width in Micrometers.
		$LOW_PAPER_WIDTH_A4 = 21006, _ ; A4 paper width in Micrometers.
		$LOW_PAPER_WIDTH_A3 = 29693, _ ; A3 paper width in Micrometers.
		$LOW_PAPER_WIDTH_B6ISO = 12497, _ ; B6ISO paper width in Micrometers.
		$LOW_PAPER_WIDTH_B5ISO = 17602, _ ; B5ISO paper width in Micrometers.
		$LOW_PAPER_WIDTH_B4ISO = 24994, _ ; B4ISO paper width in Micrometers.
		$LOW_PAPER_WIDTH_LETTER = 21590, _ ; Letter paper width in Micrometers.
		$LOW_PAPER_WIDTH_LEGAL = 21590, _ ; Legal paper width in Micrometers.
		$LOW_PAPER_WIDTH_LONG_BOND = 21590, _ ; Long Bond paper width in Micrometers.
		$LOW_PAPER_WIDTH_TABLOID = 27940, _ ; Tabloid paper width in Micrometers.
		$LOW_PAPER_WIDTH_B6JIS = 12801, _ ; B6JIS paper width in Micrometers.
		$LOW_PAPER_WIDTH_B5JIS = 18212, _ ; B5JIS paper width in Micrometers.
		$LOW_PAPER_WIDTH_B4JIS = 25705, _ ; B4JIS paper width in Micrometers.
		$LOW_PAPER_WIDTH_16KAI = 18390, _ ; 16KAI paper width in Micrometers.
		$LOW_PAPER_WIDTH_32KAI = 13005, _ ; 32KAI paper width in Micrometers.
		$LOW_PAPER_WIDTH_BIG_32KAI = 13995, _ ; Big 32KAI paper width in Micrometers.
		$LOW_PAPER_WIDTH_DLENVELOPE = 10998, _ ; DL Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_C6ENVELOPE = 11405, _ ; C6 Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_C6_5_ENVELOPE = 11405, _ ; C6/5 Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_C5ENVELOPE = 16205, _ ; C5 Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_C4ENVELOPE = 22911, _ ; C4 Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_6_3_4ENVELOPE = 9208, _ ; 6 3/4 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_7_3_4ENVELOPE = 9855, _ ; 7 3/4 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_9ENVELOPE = 9843, _ ; 9 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_10ENVELOPE = 10490, _ ; 10 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_11ENVELOPE = 11430, _ ; 11 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_12ENVELOPE = 12065, _ ; 12 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_JAP_POSTCARD = 10008 ; Japanese Postcard paper width in Micrometers.

Global Const _ ; Paper Height in uM
		$LOW_PAPER_HEIGHT_A6 = 14808, _ ; A6 paper height in Micrometers.
		$LOW_PAPER_HEIGHT_A5 = 21006, _ ; A5 paper height in Micrometers.
		$LOW_PAPER_HEIGHT_A4 = 29693, _ ; A4 paper height in Micrometers.
		$LOW_PAPER_HEIGHT_A3 = 42012, _ ; A3 paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B6ISO = 17602, _ ; B6ISO paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B5ISO = 24994, _ ; B5ISO paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B4ISO = 35306, _ ; B4ISO paper height in Micrometers.
		$LOW_PAPER_HEIGHT_LETTER = 27940, _ ; Letter paper height in Micrometers.
		$LOW_PAPER_HEIGHT_LEGAL = 35560, _ ; Legal paper height in Micrometers.
		$LOW_PAPER_HEIGHT_LONG_BOND = 33020, _ ; Long Bond paper height in Micrometers.
		$LOW_PAPER_HEIGHT_TABLOID = 43180, _ ; Tabloid paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B6JIS = 18200, _ ; B6JIS paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B5JIS = 25705, _ ; B5JIS paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B4JIS = 36398, _ ; B4JIS paper height in Micrometers.
		$LOW_PAPER_HEIGHT_16KAI = 26010, _ ; 16KAI paper height in Micrometers.
		$LOW_PAPER_HEIGHT_32KAI = 18390, _ ; 32KAI paper height in Micrometers.
		$LOW_PAPER_HEIGHT_BIG_32KAI = 20295, _ ; Big 32KAI paper height in Micrometers.
		$LOW_PAPER_HEIGHT_DLENVELOPE = 21996, _ ; DL Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_C6ENVELOPE = 16205, _ ; C6 Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_C6_5_ENVELOPE = 22911, _ ; C6/5 Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_C5ENVELOPE = 22911, _ ; C5 Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_C4ENVELOPE = 32410, _ ; C4 Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_6_3_4ENVELOPE = 16510, _ ; 6 3/4 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_7_3_4ENVELOPE = 19050, _ ; 7 3/4 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_9ENVELOPE = 22543, _ ; 9 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_10ENVELOPE = 24130, _ ; 10 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_11ENVELOPE = 26365, _ ; 11 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_12ENVELOPE = 27940, _ ; 12 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_JAP_POSTCARD = 14808 ; Japanese Postcard paper height in Micrometers.

Global Const _ ; Gradient Names
		$LOW_GRAD_NAME_PASTEL_BOUQUET = "Pastel Bouquet", _ ; The "Pastel Bouquet" Gradient Preset.
		$LOW_GRAD_NAME_PASTEL_DREAM = "Pastel Dream", _ ; The "Pastel Dream" Gradient Preset.
		$LOW_GRAD_NAME_BLUE_TOUCH = "Blue Touch", _ ; The "Blue Touch" Gradient Preset.
		$LOW_GRAD_NAME_BLANK_W_GRAY = "Blank with Gray", _ ; The "Blank with Gray" Gradient Preset.
		$LOW_GRAD_NAME_SPOTTED_GRAY = "Spotted Gray", _ ; The "Spotted Gray" Gradient Preset.
		$LOW_GRAD_NAME_LONDON_MIST = "London Mist", _ ; The "London Mist" Gradient Preset.
		$LOW_GRAD_NAME_TEAL_TO_BLUE = "Teal to Blue", _ ; The "Teal to Blue" Gradient Preset.
		$LOW_GRAD_NAME_MIDNIGHT = "Midnight", _ ; The "Midnight" Gradient Preset.
		$LOW_GRAD_NAME_DEEP_OCEAN = "Deep Ocean", _ ; The "Deep Ocean" Gradient Preset.
		$LOW_GRAD_NAME_SUBMARINE = "Submarine", _ ; The "Submarine" Gradient Preset.
		$LOW_GRAD_NAME_GREEN_GRASS = "Green Grass", _ ; The "Green Grass" Gradient Preset.
		$LOW_GRAD_NAME_NEON_LIGHT = "Neon Light", _ ; The "Neon Light" Gradient Preset.
		$LOW_GRAD_NAME_SUNSHINE = "Sunshine", _ ; The "Sunshine" Gradient Preset.
		$LOW_GRAD_NAME_PRESENT = "Present", _ ; The "Present" Gradient Preset.
		$LOW_GRAD_NAME_MAHOGANY = "Mahogany" ; The "Mahogany" Gradient Preset.

Global Const _ ; Page Layout
		$LOW_PAGE_LAYOUT_ALL = 0, _ ; Page style shows both odd(Right) and even(Left) pages. With left and right margins.
		$LOW_PAGE_LAYOUT_LEFT = 1, _ ; Page style shows only even(Left) pages. Odd pages are shown as blank pages. With left and right margins.
		$LOW_PAGE_LAYOUT_RIGHT = 2, _ ; Page style shows only odd(Right) pages. Even pages are shown as blank pages. With left and right margins.
		$LOW_PAGE_LAYOUT_MIRRORED = 3 ; Page style shows both odd(Right) and even(Left) pages with inner and outer margins.

Global Const _ ; Numbering Style Type
		$LOW_NUM_STYLE_CHARS_UPPER_LETTER = 0, _ ; Numbering is put in upper case letters. ("A, B, C, D)
		$LOW_NUM_STYLE_CHARS_LOWER_LETTER = 1, _ ; Numbering is in lower case letters. (a, b, c, d)
		$LOW_NUM_STYLE_ROMAN_UPPER = 2, _ ; Numbering is in Roman numbers with upper case letters. (I, II, III)
		$LOW_NUM_STYLE_ROMAN_LOWER = 3, _ ; Numbering is in Roman numbers with lower case letters. (i, ii, iii).
		$LOW_NUM_STYLE_ARABIC = 4, _ ; Numbering is in Arabic numbers. (1, 2, 3, 4),
		$LOW_NUM_STYLE_NUMBER_NONE = 5, _ ; Numbering is invisible.
		$LOW_NUM_STYLE_CHAR_SPECIAL = 6, _ ; Use a character from a specified font.
		$LOW_NUM_STYLE_PAGE_DESCRIPTOR = 7, _ ; Numbering is specified in the page style.
		$LOW_NUM_STYLE_BITMAP = 8, _ ; Numbering is displayed as a bitmap graphic.
		$LOW_NUM_STYLE_CHARS_UPPER_LETTER_N = 9, _ ; Numbering is put in upper case letters. (A, B, Y, Z, AA, BB)
		$LOW_NUM_STYLE_CHARS_LOWER_LETTER_N = 10, _ ; Numbering is put in lower case letters. (a, b, y, z, aa, bb)
		$LOW_NUM_STYLE_TRANSLITERATION = 11, _ ; A transliteration module will be used to produce numbers in Chinese, Japanese, etc.
		$LOW_NUM_STYLE_NATIVE_NUMBERING = 12, _ ; The NativeNumberSupplier service will be called to produce numbers in native languages.
		$LOW_NUM_STYLE_FULLWIDTH_ARABIC = 13, _ ; Numbering for full width Arabic number.
		$LOW_NUM_STYLE_CIRCLE_NUMBER = 14, _ ; Bullet for Circle Number.
		$LOW_NUM_STYLE_NUMBER_LOWER_ZH = 15, _ ; Numbering for Chinese lower case number.
		$LOW_NUM_STYLE_NUMBER_UPPER_ZH = 16, _ ; Numbering for Chinese upper case number.
		$LOW_NUM_STYLE_NUMBER_UPPER_ZH_TW = 17, _ ; Numbering for Traditional Chinese upper case number.
		$LOW_NUM_STYLE_TIAN_GAN_ZH = 18, _ ; Bullet for Chinese Tian Gan.
		$LOW_NUM_STYLE_DI_ZI_ZH = 19, _ ; Bullet for Chinese Di Zi.
		$LOW_NUM_STYLE_NUMBER_TRADITIONAL_JA = 20, _ ; Numbering for Japanese traditional number.
		$LOW_NUM_STYLE_AIU_FULLWIDTH_JA = 21, _ ; Bullet for Japanese AIU fullwidth.
		$LOW_NUM_STYLE_AIU_HALFWIDTH_JA = 22, _ ; Bullet for Japanese AIU halfwidth.
		$LOW_NUM_STYLE_IROHA_FULLWIDTH_JA = 23, _ ; Bullet for Japanese IROHA fullwidth.
		$LOW_NUM_STYLE_IROHA_HALFWIDTH_JA = 24, _ ; Bullet for Japanese IROHA halfwidth.
		$LOW_NUM_STYLE_NUMBER_UPPER_KO = 25, _ ; Numbering for Korean upper case number.
		$LOW_NUM_STYLE_NUMBER_HANGUL_KO = 26, _ ; Numbering for Korean Hangul number.
		$LOW_NUM_STYLE_HANGUL_JAMO_KO = 27, _ ; Bullet for Korean Hangul Jamo.
		$LOW_NUM_STYLE_HANGUL_SYLLABLE_KO = 28, _ ; Bullet for Korean Hangul Syllable.
		$LOW_NUM_STYLE_HANGUL_CIRCLED_JAMO_KO = 29, _ ; Bullet for Korean Hangul Circled Jamo.
		$LOW_NUM_STYLE_HANGUL_CIRCLED_SYLLABLE_KO = 30, _ ; Bullet for Korean Hangul Circled Syllable.
		$LOW_NUM_STYLE_CHARS_ARABIC = 31, _ ; Numbering in Arabic alphabet letters.
		$LOW_NUM_STYLE_CHARS_THAI = 32, _ ; Numbering in Thai alphabet letters.
		$LOW_NUM_STYLE_CHARS_HEBREW = 33, _ ; Numbering in Hebrew alphabet letters.
		$LOW_NUM_STYLE_CHARS_NEPALI = 34, _ ; Numbering in Nepali alphabet letters.
		$LOW_NUM_STYLE_CHARS_KHMER = 35, _ ; Numbering in Khmer alphabet letters.
		$LOW_NUM_STYLE_CHARS_LAO = 36, _ ; Numbering in Lao alphabet letters.
		$LOW_NUM_STYLE_CHARS_TIBETAN = 37, _ ; Numbering in Tibetan/Dzongkha alphabet letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_BG = 38, _ ; Numbering in Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_BG = 39, _ ; Numbering in Cyrillic alphabet lower case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_BG = 40, _ ; Numbering in Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_BG = 41, _ ; Numbering in Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_RU = 42, _ ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_RU = 43, _ ; Numbering in Russian Cyrillic alphabet lower case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_RU = 44, _ ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_RU = 45, _ ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_PERSIAN = 46, _ ; Numbering in Persian alphabet letters.
		$LOW_NUM_STYLE_CHARS_MYANMAR = 47, _ ; Numbering in Myanmar alphabet letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_SR = 48, _ ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_SR = 49, _ ; Numbering in Russian Serbian alphabet lower case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_SR = 50, _ ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_SR = 51, _ ;
		$LOW_NUM_STYLE_CHARS_GREEK_UPPER_LETTER = 52, _ ;
		$LOW_NUM_STYLE_CHARS_GREEK_LOWER_LETTER = 53, _ ;
		$LOW_NUM_STYLE_CHARS_ARABIC_ABJAD = 54, _ ;
		$LOW_NUM_STYLE_CHARS_PERSIAN_WORD = 55, _ ;
		$LOW_NUM_STYLE_NUMBER_HEBREW = 56, _ ;
		$LOW_NUM_STYLE_NUMBER_ARABIC_INDIC = 57, _ ;
		$LOW_NUM_STYLE_NUMBER_EAST_ARABIC_INDIC = 58, _ ;
		$LOW_NUM_STYLE_NUMBER_INDIC_DEVANAGARI = 59, _ ;
		$LOW_NUM_STYLE_TEXT_NUMBER = 60, _ ;
		$LOW_NUM_STYLE_TEXT_CARDINAL = 61, _ ;
		$LOW_NUM_STYLE_TEXT_ORDINAL = 62, _ ;
		$LOW_NUM_STYLE_SYMBOL_CHICAGO = 63, _ ;
		$LOW_NUM_STYLE_ARABIC_ZERO = 64, _ ;
		$LOW_NUM_STYLE_ARABIC_ZERO3 = 65, _ ;
		$LOW_NUM_STYLE_ARABIC_ZERO4 = 66, _ ;
		$LOW_NUM_STYLE_ARABIC_ZERO5 = 67, _ ;
		$LOW_NUM_STYLE_SZEKELY_ROVAS = 68, _ ;
		$LOW_NUM_STYLE_NUMBER_DIGITAL_KO = 69, _ ;
		$LOW_NUM_STYLE_NUMBER_DIGITAL2_KO = 70, _ ;
		$LOW_NUM_STYLE_NUMBER_LEGAL_KO = 71 ;

; Line Style
Global Const _
		$LOW_LINE_STYLE_NONE = 0, _
		$LOW_LINE_STYLE_SOLID = 1, _
		$LOW_LINE_STYLE_DOTTED = 2, _
		$LOW_LINE_STYLE_DASHED = 3

; Vertical Alignment
Global Const _
		$LOW_ALIGN_VERT_TOP = 0, _
		$LOW_ALIGN_VERT_MIDDLE = 1, _
		$LOW_ALIGN_VERT_BOTTOM = 2

; Horizontal Alignment
Global Const _
		$LOW_ALIGN_HORI_LEFT = 0, _
		$LOW_ALIGN_HORI_CENTER = 1, _
		$LOW_ALIGN_HORI_RIGHT = 2

; Gradient Type
Global Const _
		$LOW_GRAD_TYPE_OFF = -1, _
		$LOW_GRAD_TYPE_LINEAR = 0, _
		$LOW_GRAD_TYPE_AXIAL = 1, _
		$LOW_GRAD_TYPE_RADIAL = 2, _
		$LOW_GRAD_TYPE_ELLIPTICAL = 3, _
		$LOW_GRAD_TYPE_SQUARE = 4, _
		$LOW_GRAD_TYPE_RECT = 5

; Follow By
Global Const _
		$LOW_FOLLOW_BY_TABSTOP = 0, _
		$LOW_FOLLOW_BY_SPACE = 1, _
		$LOW_FOLLOW_BY_NOTHING = 2, _
		$LOW_FOLLOW_BY_NEWLINE = 3

; Cursor Status
Global Enum $LOW_CURSOR_STAT_IS_COLLAPSED, _
		$LOW_CURSOR_STAT_IS_START_OF_WORD, _
		$LOW_CURSOR_STAT_IS_END_OF_WORD, _
		$LOW_CURSOR_STAT_IS_START_OF_SENTENCE, _
		$LOW_CURSOR_STAT_IS_END_OF_SENTENCE, _
		$LOW_CURSOR_STAT_IS_START_OF_PAR, _
		$LOW_CURSOR_STAT_IS_END_OF_PAR, _
		$LOW_CURSOR_STAT_IS_START_OF_LINE, _
		$LOW_CURSOR_STAT_IS_END_OF_LINE, _
		$LOW_CURSOR_STAT_GET_PAGE, _
		$LOW_CURSOR_STAT_GET_RANGE_NAME

; Relative to
Global Const _
		$LOW_RELATIVE_ROW = -1, _
		$LOW_RELATIVE_PARAGRAPH = 0, _
		$LOW_RELATIVE_PARAGRAPH_TEXT = 1, _
		$LOW_RELATIVE_CHARACTER = 2, _
		$LOW_RELATIVE_PAGE_LEFT = 3, _
		$LOW_RELATIVE_PAGE_RIGHT = 4, _
		$LOW_RELATIVE_PARAGRAPH_LEFT = 5, _
		$LOW_RELATIVE_PARAGRAPH_RIGHT = 6, _
		$LOW_RELATIVE_PAGE = 7, _
		$LOW_RELATIVE_PAGE_PRINT = 8, _
		$LOW_RELATIVE_TEXT_LINE = 9, _
		$LOW_RELATIVE_PAGE_PRINT_BOTTOM = 10, _
		$LOW_RELATIVE_PAGE_PRINT_TOP = 11

; Anchor Type
Global Const _
		$LOW_ANCHOR_AT_PARAGRAPH = 0, _
		$LOW_ANCHOR_AS_CHARACTER = 1, _
		$LOW_ANCHOR_AT_PAGE = 2, _
		$LOW_ANCHOR_AT_FRAME = 3, _
		$LOW_ANCHOR_AT_CHARACTER = 4

; Wrap Type
Global Const _
		$LOW_WRAP_MODE_NONE = 0, _
		$LOW_WRAP_MODE_THROUGH = 1, _
		$LOW_WRAP_MODE_PARALLEL = 2, _
		$LOW_WRAP_MODE_DYNAMIC = 3, _
		$LOW_WRAP_MODE_LEFT = 4, _
		$LOW_WRAP_MODE_RIGHT = 5

; Text Adjust
Global Const _
		$LOW_TXT_ADJ_VERT_TOP = 0, _
		$LOW_TXT_ADJ_VERT_CENTER = 1, _
		$LOW_TXT_ADJ_VERT_BOTTOM = 2, _
		$LOW_TXT_ADJ_VERT_BLOCK = 3

; Frame Target
Global Const _
		$LOW_FRAME_TARGET_NONE = "", _
		$LOW_FRAME_TARGET_TOP = "_top", _
		$LOW_FRAME_TARGET_PARENT = "_parent", _
		$LOW_FRAME_TARGET_BLANK = "_blank", _
		$LOW_FRAME_TARGET_SELF = "_self"

; Footnote Count type
Global Const _
		$LOW_FOOTNOTE_COUNT_PER_PAGE = 0, _
		$LOW_FOOTNOTE_COUNT_PER_CHAP = 1, _
		$LOW_FOOTNOTE_COUNT_PER_DOC = 2

; Page Number Type
Global Const _
		$LOW_PAGE_NUM_TYPE_PREV = 0, _
		$LOW_PAGE_NUM_TYPE_CURRENT = 1, _
		$LOW_PAGE_NUM_TYPE_NEXT = 2

; Field Chapter Display Type
Global Const _
		$LOW_FIELD_CHAP_FRMT_NAME = 0, _
		$LOW_FIELD_CHAP_FRMT_NUMBER = 1, _
		$LOW_FIELD_CHAP_FRMT_NAME_NUMBER = 2, _
		$LOW_FIELD_CHAP_FRMT_NO_PREFIX_SUFFIX = 3, _
		$LOW_FIELD_CHAP_FRMT_DIGIT = 4

; User Data Field Type
Global Const _
		$LOW_FIELD_USER_DATA_COMPANY = 0, _
		$LOW_FIELD_USER_DATA_FIRST_NAME = 1, _
		$LOW_FIELD_USER_DATA_NAME = 2, _
		$LOW_FIELD_USER_DATA_SHORTCUT = 3, _
		$LOW_FIELD_USER_DATA_STREET = 4, _
		$LOW_FIELD_USER_DATA_COUNTRY = 5, _
		$LOW_FIELD_USER_DATA_ZIP = 6, _
		$LOW_FIELD_USER_DATA_CITY = 7, _
		$LOW_FIELD_USER_DATA_TITLE = 8, _
		$LOW_FIELD_USER_DATA_POSITION = 9, _
		$LOW_FIELD_USER_DATA_PHONE_PRIVATE = 10, _
		$LOW_FIELD_USER_DATA_PHONE_COMPANY = 11, _
		$LOW_FIELD_USER_DATA_FAX = 12, _
		$LOW_FIELD_USER_DATA_EMAIL = 13, _
		$LOW_FIELD_USER_DATA_STATE = 14

; File Name Field Type
Global Const _
		$LOW_FIELD_FILENAME_FULL_PATH = 0, _               ; The content of the URL is completely displayed.
		$LOW_FIELD_FILENAME_PATH = 1, _                    ; Only the path of the file is displayed.
		$LOW_FIELD_FILENAME_NAME = 2, _                    ; Only the name of the file without the file extension is displayed.
		$LOW_FIELD_FILENAME_NAME_AND_EXT = 3, _            ; The file name including the file extension is displayed.
		$LOW_FIELD_FILENAME_CATEGORY = 4, _
		$LOW_FIELD_FILENAME_TEMPLATE_NAME = 5

; Format Key Type
Global Const _
		$LOW_FORMAT_KEYS_ALL = 0, _
		$LOW_FORMAT_KEYS_DEFINED = 1, _
		$LOW_FORMAT_KEYS_DATE = 2, _
		$LOW_FORMAT_KEYS_TIME = 4, _
		$LOW_FORMAT_KEYS_DATE_TIME = 6, _
		$LOW_FORMAT_KEYS_CURRENCY = 8, _
		$LOW_FORMAT_KEYS_NUMBER = 16, _
		$LOW_FORMAT_KEYS_SCIENTIFIC = 32, _
		$LOW_FORMAT_KEYS_FRACTION = 64, _
		$LOW_FORMAT_KEYS_PERCENT = 128, _
		$LOW_FORMAT_KEYS_TEXT = 256, _
		$LOW_FORMAT_KEYS_LOGICAL = 1024, _
		$LOW_FORMAT_KEYS_UNDEFINED = 2048, _
		$LOW_FORMAT_KEYS_EMPTY = 4096, _
		$LOW_FORMAT_KEYS_DURATION = 8196

; Reference Field Type
Global Const _
		$LOW_FIELD_REF_TYPE_REF_MARK = 0, _
		$LOW_FIELD_REF_TYPE_SEQ_FIELD = 1, _
		$LOW_FIELD_REF_TYPE_BOOKMARK = 2, _
		$LOW_FIELD_REF_TYPE_FOOTNOTE = 3, _
		$LOW_FIELD_REF_TYPE_ENDNOTE = 4

; Type of Reference
Global Const _
		$LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED = 0, _
		$LOW_FIELD_REF_USING_CHAPTER = 1, _
		$LOW_FIELD_REF_USING_REF_TEXT = 2, _
		$LOW_FIELD_REF_USING_ABOVE_BELOW = 3, _
		$LOW_FIELD_REF_USING_PAGE_NUM_STYLED = 4, _
		$LOW_FIELD_REF_USING_CAT_AND_NUM = 5, _
		$LOW_FIELD_REF_USING_CAPTION = 6, _
		$LOW_FIELD_REF_USING_NUMBERING = 7, _
		$LOW_FIELD_REF_USING_NUMBER = 8, _
		$LOW_FIELD_REF_USING_NUMBER_NO_CONT = 9, _
		$LOW_FIELD_REF_USING_NUMBER_CONT = 10

; Count Field Type
Global Enum $LOW_FIELD_COUNT_TYPE_CHARACTERS = 0, _
		$LOW_FIELD_COUNT_TYPE_IMAGES, _
		$LOW_FIELD_COUNT_TYPE_OBJECTS, _
		$LOW_FIELD_COUNT_TYPE_PAGES, _
		$LOW_FIELD_COUNT_TYPE_PARAGRAPHS, _
		$LOW_FIELD_COUNT_TYPE_TABLES, _
		$LOW_FIELD_COUNT_TYPE_WORDS

; Regular Field Types
Global Enum Step *2 _
		$LOW_FIELD_TYPE_ALL = 1, _
		$LOW_FIELD_TYPE_COMMENT, _
		$LOW_FIELD_TYPE_AUTHOR, _
		$LOW_FIELD_TYPE_CHAPTER, _
		$LOW_FIELD_TYPE_CHAR_COUNT, _
		$LOW_FIELD_TYPE_COMBINED_CHAR, _
		$LOW_FIELD_TYPE_COND_TEXT, _
		$LOW_FIELD_TYPE_DATE_TIME, _
		$LOW_FIELD_TYPE_INPUT_LIST, _
		$LOW_FIELD_TYPE_EMB_OBJ_COUNT, _
		$LOW_FIELD_TYPE_SENDER, _
		$LOW_FIELD_TYPE_FILENAME, _
		$LOW_FIELD_TYPE_SHOW_VAR, _
		$LOW_FIELD_TYPE_INSERT_REF, _
		$LOW_FIELD_TYPE_IMAGE_COUNT, _
		$LOW_FIELD_TYPE_HIDDEN_PAR, _
		$LOW_FIELD_TYPE_HIDDEN_TEXT, _
		$LOW_FIELD_TYPE_INPUT, _
		$LOW_FIELD_TYPE_PLACEHOLDER, _
		$LOW_FIELD_TYPE_MACRO, _
		$LOW_FIELD_TYPE_PAGE_COUNT, _
		$LOW_FIELD_TYPE_PAGE_NUM, _
		$LOW_FIELD_TYPE_PAR_COUNT, _
		$LOW_FIELD_TYPE_SHOW_PAGE_VAR, _
		$LOW_FIELD_TYPE_SET_PAGE_VAR, _
		$LOW_FIELD_TYPE_SCRIPT, _
		$LOW_FIELD_TYPE_SET_VAR, _
		$LOW_FIELD_TYPE_TABLE_COUNT, _
		$LOW_FIELD_TYPE_TEMPLATE_NAME, _
		$LOW_FIELD_TYPE_URL, _
		$LOW_FIELD_TYPE_WORD_COUNT

; Advanced Field Types
Global Enum Step *2 _
		$LOW_FIELDADV_TYPE_ALL = 1, _
		$LOW_FIELDADV_TYPE_BIBLIOGRAPHY, _
		$LOW_FIELDADV_TYPE_DATABASE, _
		$LOW_FIELDADV_TYPE_DATABASE_SET_NUM, _
		$LOW_FIELDADV_TYPE_DATABASE_NAME, _
		$LOW_FIELDADV_TYPE_DATABASE_NEXT_SET, _
		$LOW_FIELDADV_TYPE_DATABASE_NAME_OF_SET, _
		$LOW_FIELDADV_TYPE_DDE, _
		$LOW_FIELDADV_TYPE_INPUT_USER, _
		$LOW_FIELDADV_TYPE_USER

; Document Information Field Types
Global Enum Step *2 _
		$LOW_FIELD_DOCINFO_TYPE_ALL = 1, _
		$LOW_FIELD_DOCINFO_TYPE_MOD_AUTH, _
		$LOW_FIELD_DOCINFO_TYPE_MOD_DATE_TIME, _
		$LOW_FIELD_DOCINFO_TYPE_CREATE_AUTH, _
		$LOW_FIELD_DOCINFO_TYPE_CREATE_DATE_TIME, _
		$LOW_FIELD_DOCINFO_TYPE_CUSTOM, _
		$LOW_FIELD_DOCINFO_TYPE_COMMENTS, _
		$LOW_FIELD_DOCINFO_TYPE_EDIT_TIME, _
		$LOW_FIELD_DOCINFO_TYPE_KEYWORDS, _
		$LOW_FIELD_DOCINFO_TYPE_PRINT_AUTH, _
		$LOW_FIELD_DOCINFO_TYPE_PRINT_DATE_TIME, _
		$LOW_FIELD_DOCINFO_TYPE_REVISION, _
		$LOW_FIELD_DOCINFO_TYPE_SUBJECT, _
		$LOW_FIELD_DOCINFO_TYPE_TITLE

; Placeholder Type
Global Const _
		$LOW_FIELD_PLACEHOLD_TYPE_TEXT = 0, _
		$LOW_FIELD_PLACEHOLD_TYPE_TABLE = 1, _
		$LOW_FIELD_PLACEHOLD_TYPE_FRAME = 2, _
		$LOW_FIELD_PLACEHOLD_TYPE_GRAPHIC = 3, _
		$LOW_FIELD_PLACEHOLD_TYPE_OBJECT = 4