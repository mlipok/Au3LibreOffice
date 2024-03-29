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
		$LOC_CELL_ALIGN_HORI_DEFAULT = 0, _                     ; The default alignment is used (left for numbers, right for text).
		$LOC_CELL_ALIGN_HORI_LEFT = 1, _                        ; The contents are printed from left to right.
		$LOC_CELL_ALIGN_HORI_CENTER = 2, _                      ; The contents are horizontally centered.
		$LOC_CELL_ALIGN_HORI_RIGHT = 3, _                       ; The contents are aligned to the right edge of the cell.
		$LOC_CELL_ALIGN_HORI_JUSTIFIED = 4, _                   ; The contents are justified to the cell width.
		$LOC_CELL_ALIGN_HORI_FILLED = 5, _                      ; The contents are repeated to fill the cell.
		$LOC_CELL_ALIGN_HORI_DISTRIBUTED = 6                    ; The contents are evenly aligned across the whole cell. Unlike Justified, it justifies the very last line of text, too.

; Cell Content Vertical Alignment
Global Const _
		$LOC_CELL_ALIGN_VERT_DEFAULT = 0, _                     ; The default alignment is used.
		$LOC_CELL_ALIGN_VERT_TOP = 1, _                         ; The contents are aligned with the upper edge of the cell.
		$LOC_CELL_ALIGN_VERT_MIDDLE = 2, _                      ; The contents are aligned to the vertical middle of the cell.
		$LOC_CELL_ALIGN_VERT_BOTTOM = 3, _                      ; The contents are aligned to the lower edge of the cell.
		$LOC_CELL_ALIGN_VERT_JUSTIFIED = 4, _                   ; The contents are justified to the cell height.
		$LOC_CELL_ALIGN_VERT_DISTRIBUTED = 5                    ; The same as Justified, unless the text orientation is vertical. Then it behaves similarly to the horizontal Distributed setting, i.e. the very last line is justified, too.

; Cell Delete Mode Constants
Global Const _
		$LOC_CELL_DELETE_MODE_NONE = 0, _                       ; No cells are moved -- Nothing happens.
		$LOC_CELL_DELETE_MODE_UP = 1, _                         ; The cells below the inserted Cells are moved up.
		$LOC_CELL_DELETE_MODE_LEFT = 2, _                       ; The cells to the right of the inserted cells are moved left.
		$LOC_CELL_DELETE_MODE_ROWS = 3, _                       ; Entire rows below the inserted cells are moved up.
		$LOC_CELL_DELETE_MODE_COLUMNS = 4                       ; Entire columns to the right of the inserted cells are moved left.

; Cell Content Type Flag Constants
Global Const _
		$LOC_CELL_FLAG_VALUE = 1, _                             ; Cell Contents that are numeric values but are not formatted as dates or times.
		$LOC_CELL_FLAG_DATE_TIME = 2, _                         ; Cell Contents that are numeric values that have a date or time number format.
		$LOC_CELL_FLAG_STRING = 4, _                            ; Cell Contents that are Strings.
		$LOC_CELL_FLAG_ANNOTATION = 8, _                        ; Cell Annotations
		$LOC_CELL_FLAG_FORMULA = 16, _                          ; Cell Contents that are Formulas.
		$LOC_CELL_FLAG_HARD_ATTR = 32, _                        ; Cells with explicit formatting, but not the formatting which is applied implicitly through styles.
		$LOC_CELL_FLAG_STYLES = 64, _                           ; Cells with Styles applied.
		$LOC_CELL_FLAG_OBJECTS = 128, _                         ; Cell Contents that are Drawing Objects.
		$LOC_CELL_FLAG_EDIT_ATTR = 256, _                       ; Cells containing formatting within parts of the cell contents.
		$LOC_CELL_FLAG_FORMATTED = 512, _                       ; Cells with formatting within the cells or cells with more than one paragraph within the cells.
		$LOC_CELL_FLAG_ALL = 1023                               ; All flags listed above.

; Cell Insert Mode Constants
Global Const _
		$LOC_CELL_INSERT_MODE_NONE = 0, _                       ; No cells are moved -- Nothing happens.
		$LOC_CELL_INSERT_MODE_DOWN = 1, _                       ; The cells below the inserted Cells are moved down.
		$LOC_CELL_INSERT_MODE_RIGHT = 2, _                      ; The cells to the right of the inserted cells are moved right.
		$LOC_CELL_INSERT_MODE_ROWS = 3, _                       ; Entire rows below the inserted cells are moved down.
		$LOC_CELL_INSERT_MODE_COLUMNS = 4                       ; Entire columns to the right of the inserted cells are moved right.

; Cell Content Rotation Reference
Global Const _
		$LOC_CELL_ROTATE_REF_LOWER_CELL_BORDER = 0, _           ; Writes the rotated text from the bottom cell edge outwards.
		$LOC_CELL_ROTATE_REF_UPPER_CELL_BORDER = 1, _           ; Writes the rotated text from the top cell edge outwards.
		$LOC_CELL_ROTATE_REF_INSIDE_CELLS = 3                   ; Writes the rotated text only within the cell.

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

; Compute Functions
Global Const _
		$LOC_COMPUTE_NONE = 0, _                                ; Nothing is calculated.
		$LOC_COMPUTE_AUTO = 1, _                                ; Uses SUM if all values in the range are numbers, else uses COUNT.
		$LOC_COMPUTE_SUM = 2, _                                 ; Adds all numerical values in the Range.
		$LOC_COMPUTE_COUNT = 3, _                               ; Count all cells containing a string or value.
		$LOC_COMPUTE_AVERAGE = 4, _                             ; Average all numerical values in a range.
		$LOC_COMPUTE_MAX = 5, _                                 ; Find the maximum numerical value in the range.
		$LOC_COMPUTE_MIN = 6, _                                 ; Find the minimum numerical value in the range.
		$LOC_COMPUTE_PRODUCT = 7, _                             ; The result of multiplying of all numbers in the range.
		$LOC_COMPUTE_COUNTNUMS = 8, _                           ; Count the number of cells containing numerical values in the range.
		$LOC_COMPUTE_STDEV = 9, _                               ; Standard deviation based on a sample.
		$LOC_COMPUTE_STDEVP = 10, _                             ; Standard deviation based on the entire population.
		$LOC_COMPUTE_VAR = 11, _                                ; Variance based on a sample.
		$LOC_COMPUTE_VARP = 12                                  ; Variance based on the entire population.

; Cursor Type Related Constants
Global Const _
		$LOC_CURTYPE_TEXT_CURSOR = 1, _                         ; Cursor is a Text Cursor type.
		$LOC_CURTYPE_SHEET_CURSOR = 2, _                        ; Cursor is a Sheet Cursor type.
		$LOC_CURTYPE_PARAGRAPH = 3, _                           ; Object is a Paragraph Object.
		$LOC_CURTYPE_TEXT_PORTION = 4                           ; Object is a Paragraph Text Portion Object.

; Fill Date Mode
Global Const _
		$LOC_FILL_DATE_MODE_DAY = 0, _                          ; For each Cell a single day is added.
		$LOC_FILL_DATE_MODE_WEEKDAY = 1, _                      ; For each Cell a single day is added, skipping weekends.
		$LOC_FILL_DATE_MODE_MONTH = 2, _                        ; For each Cell one month is added without modifiying the day.
		$LOC_FILL_DATE_MODE_YEAR = 3                            ; For each Cell a year is added without modifiying the day or month.

; Fill Direction
Global Const _
		$LOC_FILL_DIR_DOWN = 0, _                               ; Rows are filled from top to bottom.
		$LOC_FILL_DIR_RIGHT = 1, _                              ; Columns are filled from left to right.
		$LOC_FILL_DIR_TOP = 2, _                                ; Rows are filled from bottom to top.
		$LOC_FILL_DIR_LEFT = 3                                  ; Columns are filled from right to left.

; Fill Series Mode
Global Const _
		$LOC_FILL_MODE_SIMPLE = 0, _                            ; All cells are filled with the same value.
		$LOC_FILL_MODE_LINEAR = 1, _                            ; The initial value is increased by a specified value, per each cell processed.
		$LOC_FILL_MODE_GROWTH = 2, _                            ; The initial value is multiplied by a specified value, per each cell processed.
		$LOC_FILL_MODE_DATE = 3, _                              ; Any date the Cells is increased by the specified number of days/
		$LOC_FILL_MODE_AUTO = 4                                 ; The cells are filled using a user-defined series.

; Filter Conditions
Global Const _
		$LOC_FILTER_CONDITION_EMPTY = 0, _                      ; Show only Empty cells.
		$LOC_FILTER_CONDITION_NOT_EMPTY = 1, _                  ; Show only non-empty cells.
		$LOC_FILTER_CONDITION_EQUAL = 2, _                      ; Show only cells equal to the value set.
		$LOC_FILTER_CONDITION_NOT_EQUAL = 3, _                  ; Show only cells NOT equal to the value set.
		$LOC_FILTER_CONDITION_GREATER = 4, _                    ; Show only cells greater than the value set.
		$LOC_FILTER_CONDITION_GREATER_EQUAL = 5, _              ; Show only cells greater than or equal to the value set.
		$LOC_FILTER_CONDITION_LESS = 6, _                       ; Show only cells less than the value set.
		$LOC_FILTER_CONDITION_LESS_EQUAL = 7, _                 ; Show only cells less than or equal to the value set.
		$LOC_FILTER_CONDITION_TOP_VALUES = 8, _                 ; Show a specified number of the largest values contained in the range.
		$LOC_FILTER_CONDITION_TOP_PERCENT = 9, _                ; Show a specified percentage of the largest values contained in the range.
		$LOC_FILTER_CONDITION_BOTTOM_VALUES = 10, _             ; Show a specified number of the lowest values contained in the range.
		$LOC_FILTER_CONDITION_BOTTOM_PERCENT = 11, _            ; Show a specified percentage of the lowest values contained in the range.
		$LOC_FILTER_CONDITION_CONTAINS = 12, _                  ; Show only cells containing the specified entry.
		$LOC_FILTER_CONDITION_DOES_NOT_CONTAIN = 13, _          ; Show only cells that do not contain the specified entry.
		$LOC_FILTER_CONDITION_BEGINS_WITH = 14, _               ; Show only cells beginning with the specified entry.
		$LOC_FILTER_CONDITION_DOES_NOT_BEGIN_WITH = 15, _       ; Show only cells not beginning with the specified entry.
		$LOC_FILTER_CONDITION_ENDS_WITH = 16, _                 ; Show only cells ending with the specified entry.
		$LOC_FILTER_CONDITION_DOES_NOT_END_WITH = 17            ; Show only cells not ending with the specified entry.

; Filter Operators
Global Const _
		$LOC_FILTER_OPERATOR_AND = 0, _                         ; Both conditions have to be fulfilled.
		$LOC_FILTER_OPERATOR_OR = 1                             ; At least one of the conditions has to be fulfilled.

; Format Key Type
Global Const _
		$LOC_FORMAT_KEYS_ALL = 0, _                             ; Returns All number formats.
		$LOC_FORMAT_KEYS_DEFINED = 1, _                         ; Returns Only user-defined number formats.
		$LOC_FORMAT_KEYS_DATE = 2, _                            ; Returns Date formats.
		$LOC_FORMAT_KEYS_TIME = 4, _                            ; Returns Time formats.
		$LOC_FORMAT_KEYS_DATE_TIME = 6, _                       ; Returns Number formats which contain date and time.
		$LOC_FORMAT_KEYS_CURRENCY = 8, _                        ; Returns Currency formats.
		$LOC_FORMAT_KEYS_NUMBER = 16, _                         ; Returns Decimal number formats.
		$LOC_FORMAT_KEYS_SCIENTIFIC = 32, _                     ; Returns Scientific number formats.
		$LOC_FORMAT_KEYS_FRACTION = 64, _                       ; Returns Number formats for fractions.
		$LOC_FORMAT_KEYS_PERCENT = 128, _                       ; Returns Percentage number formats.
		$LOC_FORMAT_KEYS_TEXT = 256, _                          ; Returns Text number formats.
		$LOC_FORMAT_KEYS_LOGICAL = 1024, _                      ; Returns Boolean number formats.
		$LOC_FORMAT_KEYS_UNDEFINED = 2048, _                    ; Returns Is used as a return value if no format exists.
		$LOC_FORMAT_KEYS_EMPTY = 4096, _                        ; Returns Empty Number formats (?)
		$LOC_FORMAT_KEYS_DURATION = 8196                        ; Returns Duration number formats.

; Formula Result Type Constants
Global Const _
		$LOC_FORMULA_RESULT_TYPE_VALUE = 1, _                   ; The formula's result is a number.
		$LOC_FORMULA_RESULT_TYPE_STRING = 2, _                  ; The formula's result is a string.
		$LOC_FORMULA_RESULT_TYPE_ERROR = 4, _                   ; The formula has an error of some form.
		$LOC_FORMULA_RESULT_TYPE_ALL = 7                        ; All of the above types.

; Numbering Style Type
Global Const _
		$LOC_NUM_STYLE_CHARS_UPPER_LETTER = 0, _                ; Numbering is put in upper case letters. ("A, B, C, D)
		$LOC_NUM_STYLE_CHARS_LOWER_LETTER = 1, _                ; Numbering is in lower case letters. (a, b, c, d)
		$LOC_NUM_STYLE_ROMAN_UPPER = 2, _                       ; Numbering is in Roman numbers with upper case letters. (I, II, III)
		$LOC_NUM_STYLE_ROMAN_LOWER = 3, _                       ; Numbering is in Roman numbers with lower case letters. (i, ii, iii).
		$LOC_NUM_STYLE_ARABIC = 4, _                            ; Numbering is in Arabic numbers. (1, 2, 3, 4),
		$LOC_NUM_STYLE_NUMBER_NONE = 5, _                       ; Numbering is invisible.
		$LOC_NUM_STYLE_CHAR_SPECIAL = 6, _                      ; Use a character from a specified font.
		$LOC_NUM_STYLE_PAGE_DESCRIPTOR = 7, _                   ; Numbering is specified in the page style.
		$LOC_NUM_STYLE_BITMAP = 8, _                            ; Numbering is displayed as a bitmap graphic.
		$LOC_NUM_STYLE_CHARS_UPPER_LETTER_N = 9, _              ; Numbering is put in upper case letters. (A, B, Y, Z, AA, BB)
		$LOC_NUM_STYLE_CHARS_LOWER_LETTER_N = 10, _             ; Numbering is put in lower case letters. (a, b, y, z, aa, bb)
		$LOC_NUM_STYLE_TRANSLITERATION = 11, _                  ; A transliteration module will be used to produce numbers in Chinese, Japanese, etc.
		$LOC_NUM_STYLE_NATIVE_NUMBERING = 12, _                 ; The NativeNumberSupplier service will be called to produce numbers in native languages.
		$LOC_NUM_STYLE_FULLWIDTH_ARABIC = 13, _                 ; Numbering for full width Arabic number.
		$LOC_NUM_STYLE_CIRCLE_NUMBER = 14, _                    ; Bullet for Circle Number.
		$LOC_NUM_STYLE_NUMBER_LOWER_ZH = 15, _                  ; Numbering for Chinese lower case number.
		$LOC_NUM_STYLE_NUMBER_UPPER_ZH = 16, _                  ; Numbering for Chinese upper case number.
		$LOC_NUM_STYLE_NUMBER_UPPER_ZH_TW = 17, _               ; Numbering for Traditional Chinese upper case number.
		$LOC_NUM_STYLE_TIAN_GAN_ZH = 18, _                      ; Bullet for Chinese Tian Gan.
		$LOC_NUM_STYLE_DI_ZI_ZH = 19, _                         ; Bullet for Chinese Di Zi.
		$LOC_NUM_STYLE_NUMBER_TRADITIONAL_JA = 20, _            ; Numbering for Japanese traditional number.
		$LOC_NUM_STYLE_AIU_FULLWIDTH_JA = 21, _                 ; Bullet for Japanese AIU fullwidth.
		$LOC_NUM_STYLE_AIU_HALFWIDTH_JA = 22, _                 ; Bullet for Japanese AIU halfwidth.
		$LOC_NUM_STYLE_IROHA_FULLWIDTH_JA = 23, _               ; Bullet for Japanese IROHA fullwidth.
		$LOC_NUM_STYLE_IROHA_HALFWIDTH_JA = 24, _               ; Bullet for Japanese IROHA halfwidth.
		$LOC_NUM_STYLE_NUMBER_UPPER_KO = 25, _                  ; Numbering for Korean upper case number.
		$LOC_NUM_STYLE_NUMBER_HANGUL_KO = 26, _                 ; Numbering for Korean Hangul number.
		$LOC_NUM_STYLE_HANGUL_JAMO_KO = 27, _                   ; Bullet for Korean Hangul Jamo.
		$LOC_NUM_STYLE_HANGUL_SYLLABLE_KO = 28, _               ; Bullet for Korean Hangul Syllable.
		$LOC_NUM_STYLE_HANGUL_CIRCLED_JAMO_KO = 29, _           ; Bullet for Korean Hangul Circled Jamo.
		$LOC_NUM_STYLE_HANGUL_CIRCLED_SYLLABLE_KO = 30, _       ; Bullet for Korean Hangul Circled Syllable.
		$LOC_NUM_STYLE_CHARS_ARABIC = 31, _                     ; Numbering in Arabic alphabet letters.
		$LOC_NUM_STYLE_CHARS_THAI = 32, _                       ; Numbering in Thai alphabet letters.
		$LOC_NUM_STYLE_CHARS_HEBREW = 33, _                     ; Numbering in Hebrew alphabet letters.
		$LOC_NUM_STYLE_CHARS_NEPALI = 34, _                     ; Numbering in Nepali alphabet letters.
		$LOC_NUM_STYLE_CHARS_KHMER = 35, _                      ; Numbering in Khmer alphabet letters.
		$LOC_NUM_STYLE_CHARS_LAO = 36, _                        ; Numbering in Lao alphabet letters.
		$LOC_NUM_STYLE_CHARS_TIBETAN = 37, _                    ; Numbering in Tibetan/Dzongkha alphabet letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_BG = 38, _   ; Numbering in Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_BG = 39, _   ; Numbering in Cyrillic alphabet lower case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_BG = 40, _ ; Numbering in Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_BG = 41, _ ; Numbering in Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_RU = 42, _   ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_RU = 43, _   ; Numbering in Russian Cyrillic alphabet lower case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_RU = 44, _ ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_RU = 45, _ ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_PERSIAN = 46, _                    ; Numbering in Persian alphabet letters.
		$LOC_NUM_STYLE_CHARS_MYANMAR = 47, _                    ; Numbering in Myanmar alphabet letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_SR = 48, _   ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_SR = 49, _   ; Numbering in Russian Serbian alphabet lower case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_SR = 50, _ ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_SR = 51, _ ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_GREEK_UPPER_LETTER = 52, _         ; Numbering in Greek alphabet upper case letters.
		$LOC_NUM_STYLE_CHARS_GREEK_LOWER_LETTER = 53, _         ; Numbering in Greek alphabet lower case letters.
		$LOC_NUM_STYLE_CHARS_ARABIC_ABJAD = 54, _               ; Numbering in Arabic alphabet using abjad sequence.
		$LOC_NUM_STYLE_CHARS_PERSIAN_WORD = 55, _               ; Numbering in Persian words.
		$LOC_NUM_STYLE_NUMBER_HEBREW = 56, _                    ; Numbering in Hebrew numerals.
		$LOC_NUM_STYLE_NUMBER_ARABIC_INDIC = 57, _              ; Numbering in Arabic-Indic numerals.
		$LOC_NUM_STYLE_NUMBER_EAST_ARABIC_INDIC = 58, _         ; Numbering in East Arabic-Indic numerals.
		$LOC_NUM_STYLE_NUMBER_INDIC_DEVANAGARI = 59, _          ; Numbering in Indic Devanagari numerals.
		$LOC_NUM_STYLE_TEXT_NUMBER = 60, _                      ; Numbering in ordinal numbers of the language of the text node. (1st, 2nd, 3rd)
		$LOC_NUM_STYLE_TEXT_CARDINAL = 61, _                    ; Numbering in cardinal numbers of the language of the text node. (One, Two)
		$LOC_NUM_STYLE_TEXT_ORDINAL = 62, _                     ; Numbering in ordinal numbers of the language of the text node. (First, Second)
		$LOC_NUM_STYLE_SYMBOL_CHICAGO = 63, _                   ; Footnoting symbols according the University of Chicago style.
		$LOC_NUM_STYLE_ARABIC_ZERO = 64, _                      ; Numbering is in Arabic numbers, padded with zero to have a length of at least two. (01, 02)
		$LOC_NUM_STYLE_ARABIC_ZERO3 = 65, _                     ; Numbering is in Arabic numbers, padded with zero to have a length of at least three.
		$LOC_NUM_STYLE_ARABIC_ZERO4 = 66, _                     ; Numbering is in Arabic numbers, padded with zero to have a length of at least four.
		$LOC_NUM_STYLE_ARABIC_ZERO5 = 67, _                     ; Numbering is in Arabic numbers, padded with zero to have a length of at least five.
		$LOC_NUM_STYLE_SZEKELY_ROVAS = 68, _                    ; Numbering is in Szekely rovas (Old Hungarian) numerals.
		$LOC_NUM_STYLE_NUMBER_DIGITAL_KO = 69, _                ; Numbering is in Korean Digital number.
		$LOC_NUM_STYLE_NUMBER_DIGITAL2_KO = 70, _               ; Numbering is in Korean Digital Number, reserved "koreanDigital2".
		$LOC_NUM_STYLE_NUMBER_LEGAL_KO = 71                     ; Numbering is in Korean Legal Number, reserved "koreanLegal".

; Page Layout
Global Const _
		$LOC_PAGE_LAYOUT_ALL = 0, _                             ; Page style shows both odd(Right) and even(Left) pages. With left and right margins.
		$LOC_PAGE_LAYOUT_LEFT = 1, _                            ; Page style shows only even(Left) pages. Odd pages are shown as blank pages. With left and right margins.
		$LOC_PAGE_LAYOUT_RIGHT = 2, _                           ; Page style shows only odd(Right) pages. Even pages are shown as blank pages. With left and right margins.
		$LOC_PAGE_LAYOUT_MIRRORED = 3                           ; Page style shows both odd(Right) and even(Left) pages with inner and outer margins.

; Paper Height in uM
Global Const _
		$LOC_PAPER_HEIGHT_A6 = 14808, _                         ; A6 paper height in Micrometers.
		$LOC_PAPER_HEIGHT_A5 = 21006, _                         ; A5 paper height in Micrometers.
		$LOC_PAPER_HEIGHT_A4 = 29693, _                         ; A4 paper height in Micrometers.
		$LOC_PAPER_HEIGHT_A3 = 42012, _                         ; A3 paper height in Micrometers.
		$LOC_PAPER_HEIGHT_B6ISO = 17602, _                      ; B6ISO paper height in Micrometers.
		$LOC_PAPER_HEIGHT_B5ISO = 24994, _                      ; B5ISO paper height in Micrometers.
		$LOC_PAPER_HEIGHT_B4ISO = 35306, _                      ; B4ISO paper height in Micrometers.
		$LOC_PAPER_HEIGHT_LETTER = 27940, _                     ; Letter paper height in Micrometers.
		$LOC_PAPER_HEIGHT_LEGAL = 35560, _                      ; Legal paper height in Micrometers.
		$LOC_PAPER_HEIGHT_LONG_BOND = 33020, _                  ; Long Bond paper height in Micrometers.
		$LOC_PAPER_HEIGHT_TABLOID = 43180, _                    ; Tabloid paper height in Micrometers.
		$LOC_PAPER_HEIGHT_B6JIS = 18200, _                      ; B6JIS paper height in Micrometers.
		$LOC_PAPER_HEIGHT_B5JIS = 25705, _                      ; B5JIS paper height in Micrometers.
		$LOC_PAPER_HEIGHT_B4JIS = 36398, _                      ; B4JIS paper height in Micrometers.
		$LOC_PAPER_HEIGHT_16KAI = 26010, _                      ; 16KAI paper height in Micrometers.
		$LOC_PAPER_HEIGHT_32KAI = 18390, _                      ; 32KAI paper height in Micrometers.
		$LOC_PAPER_HEIGHT_BIG_32KAI = 20295, _                  ; Big 32KAI paper height in Micrometers.
		$LOC_PAPER_HEIGHT_DLENVELOPE = 21996, _                 ; DL Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_C6ENVELOPE = 16205, _                 ; C6 Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_C6_5_ENVELOPE = 22911, _              ; C6/5 Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_C5ENVELOPE = 22911, _                 ; C5 Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_C4ENVELOPE = 32410, _                 ; C4 Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_6_3_4ENVELOPE = 16510, _              ; 6 3/4 Pound Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_7_3_4ENVELOPE = 19050, _              ; 7 3/4 Pound Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_9ENVELOPE = 22543, _                  ; 9 Pound Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_10ENVELOPE = 24130, _                 ; 10 Pound Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_11ENVELOPE = 26365, _                 ; 11 Pound Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_12ENVELOPE = 27940, _                 ; 12 Pound Envelope paper height in Micrometers.
		$LOC_PAPER_HEIGHT_JAP_POSTCARD = 14808                  ; Japanese Postcard paper height in Micrometers.

; Paper Width in uM
Global Const _
		$LOC_PAPER_WIDTH_A6 = 10490, _                          ; A6 paper width in Micrometers.
		$LOC_PAPER_WIDTH_A5 = 14808, _                          ; A5 paper width in Micrometers.
		$LOC_PAPER_WIDTH_A4 = 21006, _                          ; A4 paper width in Micrometers.
		$LOC_PAPER_WIDTH_A3 = 29693, _                          ; A3 paper width in Micrometers.
		$LOC_PAPER_WIDTH_B6ISO = 12497, _                       ; B6ISO paper width in Micrometers.
		$LOC_PAPER_WIDTH_B5ISO = 17602, _                       ; B5ISO paper width in Micrometers.
		$LOC_PAPER_WIDTH_B4ISO = 24994, _                       ; B4ISO paper width in Micrometers.
		$LOC_PAPER_WIDTH_LETTER = 21590, _                      ; Letter paper width in Micrometers.
		$LOC_PAPER_WIDTH_LEGAL = 21590, _                       ; Legal paper width in Micrometers.
		$LOC_PAPER_WIDTH_LONG_BOND = 21590, _                   ; Long Bond paper width in Micrometers.
		$LOC_PAPER_WIDTH_TABLOID = 27940, _                     ; Tabloid paper width in Micrometers.
		$LOC_PAPER_WIDTH_B6JIS = 12801, _                       ; B6JIS paper width in Micrometers.
		$LOC_PAPER_WIDTH_B5JIS = 18212, _                       ; B5JIS paper width in Micrometers.
		$LOC_PAPER_WIDTH_B4JIS = 25705, _                       ; B4JIS paper width in Micrometers.
		$LOC_PAPER_WIDTH_16KAI = 18390, _                       ; 16KAI paper width in Micrometers.
		$LOC_PAPER_WIDTH_32KAI = 13005, _                       ; 32KAI paper width in Micrometers.
		$LOC_PAPER_WIDTH_BIG_32KAI = 13995, _                   ; Big 32KAI paper width in Micrometers.
		$LOC_PAPER_WIDTH_DLENVELOPE = 10998, _                  ; DL Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_C6ENVELOPE = 11405, _                  ; C6 Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_C6_5_ENVELOPE = 11405, _               ; C6/5 Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_C5ENVELOPE = 16205, _                  ; C5 Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_C4ENVELOPE = 22911, _                  ; C4 Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_6_3_4ENVELOPE = 9208, _                ; 6 3/4 Pound Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_7_3_4ENVELOPE = 9855, _                ; 7 3/4 Pound Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_9ENVELOPE = 9843, _                    ; 9 Pound Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_10ENVELOPE = 10490, _                  ; 10 Pound Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_11ENVELOPE = 11430, _                  ; 11 Pound Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_12ENVELOPE = 12065, _                  ; 12 Pound Envelope paper width in Micrometers.
		$LOC_PAPER_WIDTH_JAP_POSTCARD = 10008                   ; Japanese Postcard paper width in Micrometers.

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

; Page Print Scale Mode
Global Enum _
		$LOC_SCALE_REDUCE_ENLARGE = 1, _                        ; Specifies a scaling factor to scale all printed pages.
		$LOC_SCALE_FIT_WIDTH_HEIGHT, _                          ; Specifies the maximum number of pages horizontally (width) and vertically (height) on which every sheet is to be printed.
		$LOC_SCALE_FIT_PAGES                                    ; Specifies the maximum number of pages on which every sheet is to be printed. The scale will be reduced as necessary to fit the defined number of pages.

; Search In
Global Const _
		$LOC_SEARCH_IN_FORMULAS = 0, _                          ; Searches for the search string in formulas and in non-calculated values.
		$LOC_SEARCH_IN_VALUES = 1, _                            ; Searches for the search string in values and in formula results.
		$LOC_SEARCH_IN_COMMENTS = 2                             ; Searches for the search string in comments that are attached to the cells.

; Shadow Location
Global Const _
		$LOC_SHADOW_NONE = 0, _                                 ; No shadow is applied.
		$LOC_SHADOW_TOP_LEFT = 1, _                             ; Shadow is located along the upper and left sides.
		$LOC_SHADOW_TOP_RIGHT = 2, _                            ; Shadow is located along the upper and right sides.
		$LOC_SHADOW_BOTTOM_LEFT = 3, _                          ; Shadow is located along the lower and left sides.
		$LOC_SHADOW_BOTTOM_RIGHT = 4                            ; Shadow is located along the lower and right sides.

; Sheet Link Mode
Global Const _
		$LOC_SHEET_LINK_MODE_NONE = 0, _                        ; The Sheet is not linked.
		$LOC_SHEET_LINK_MODE_NORMAL = 1, _                      ; All the Sheet's contents are copied, both values and formulas.
		$LOC_SHEET_LINK_MODE_VALUE = 2                          ; Only the Sheet's values and formula results are copied.

; Sheet Cursor Movement Constants.
Global Enum _
		$LOC_SHEETCUR_COLLAPSE_TO_CURRENT_ARRAY, _              ; Collapses or Expands the range to contain the current array formula.
		$LOC_SHEETCUR_COLLAPSE_TO_CURRENT_REGION, _             ; Collapses or Expands the range to contain all contiguous nonempty cells.
		$LOC_SHEETCUR_COLLAPSE_TO_MERGED_AREA, _                ; Collapses or Expand the range to contain merged cells that intersect the range.
		$LOC_SHEETCUR_COLLAPSE_TO_SIZE, _                       ; Beginning with the upper-left corner or the current range, set the cursor range size.
		$LOC_SHEETCUR_EXPAND_TO_ENTIRE_COLUMN, _                ; Expands the range to contain all columns that intersect the range.
		$LOC_SHEETCUR_EXPAND_TO_ENTIRE_ROW, _                   ; Expands the range to contain all rows that intersect the range.
		$LOC_SHEETCUR_GOTO_OFFSET, _                            ; Shift the cursor’s range relative to the current position. Negative numbers shift left/up; positive numbers shift right/down.
		$LOC_SHEETCUR_GOTO_START, _                             ; Move the cursor to the first filled cell at the beginning of a contiguous series of filled cells. This cell may be outside the cursor’s range.
		$LOC_SHEETCUR_GOTO_END, _                               ; Move the cursor to the last filled cell at the end of a contiguous series of filled cells. This cell may be outside the cursor’s range.
		$LOC_SHEETCUR_GOTO_NEXT, _                              ; Move the cursor to the next (right) unprotected cell.
		$LOC_SHEETCUR_GOTO_PREV, _                              ; Move the cursor to the previous (left) unprotected cell.
		$LOC_SHEETCUR_GOTO_USED_AREA_START, _                   ; Set the cursor to the start of the used area.
		$LOC_SHEETCUR_GOTO_USED_AREA_END                        ; Set the cursor to the end of the used area.

; Strikeout
Global Const _
		$LOC_STRIKEOUT_NONE = 0, _                              ; No strike out.
		$LOC_STRIKEOUT_SINGLE = 1, _                            ; Strike out the characters with a single line.
		$LOC_STRIKEOUT_DOUBLE = 2, _                            ; Strike out the characters with a double line.
		$LOC_STRIKEOUT_DONT_KNOW = 3, _                         ; The strikeout mode is not specified. For Read Only.
		$LOC_STRIKEOUT_BOLD = 4, _                              ; Strike out the characters with a bold line.
		$LOC_STRIKEOUT_SLASH = 5, _                             ; Strike out the characters with slashes.
		$LOC_STRIKEOUT_X = 6                                    ; Strike out the characters with X's.

; Text Cursor Movement Constants.
Global Enum _
		$LOC_TEXTCUR_COLLAPSE_TO_START, _                       ; Collapses the current selection to the start of the selection.
		$LOC_TEXTCUR_COLLAPSE_TO_END, _                         ; Collapses the current selection the to end of the selection.
		$LOC_TEXTCUR_GO_LEFT, _                                 ; Move the cursor left by n characters.
		$LOC_TEXTCUR_GO_RIGHT, _                                ; Move the cursor right by n characters.
		$LOC_TEXTCUR_GOTO_START, _                              ; Move the cursor to the start of the text.
		$LOC_TEXTCUR_GOTO_END                                   ; Move the cursor to the end of the text.

; Text Direction
Global Const _
		$LOC_TXT_DIR_LR = 0, _                                  ; Text within lines is written left-to-right. Typically, this is the writing mode for normal "alphabetic" text.
		$LOC_TXT_DIR_RL = 1, _                                  ; Text within a line are written right-to-left. Typically, this writing mode is used in Arabic and Hebrew text.
		$LOC_TXT_DIR_CONTEXT = 4                                ; Obtain actual writing mode from the context of the object.

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
