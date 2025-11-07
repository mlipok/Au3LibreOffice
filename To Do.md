# UDF To Do

The following is a basic list of things needing done, added to, or looked at, in this UDF.

## UDF

Things pertaining to the **entire UDF**, or **equally to all sub-Components**.

- When MultiColor Gradients (MCGR) are implemented in the UI, (Writer, Calc, etc?) I will need to modify how I made gradients work with Color stops, because there could be multiple stops, not just 2.
- Add Hatch Background option
- Add Macros functions? Plus Integrate with all FUnctions that can activate a Macro?
- Is it important to use "Get" / "set"? e.g. getText, etc? Instead of Text()?
	- <https://ask.libreoffice.org/t/calc-named-range-err-508/> See comment by JohnSUN
	- OOME 4.1 pg 306

## Writer

Things pertaining to **Writer**.

- This note is added to Form Controls: Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, <https://bugs.documentfoundation.org/show_bug.cgi?id=131196>

## Calc

Things pertaining to **Calc**.

- Find Optimal Width/Height "Add" property? Format>Rows/Columns>OptimalWidth/Height
- Add Shapes?
- Make easier way to delete header fields after reinserting header -- or way to re-identify it?

## Base

Things pertaining to **Base**.

## Impress

Things pertaining to **Impress**.

- Begin development on Impress
- For future reference:  The PresentationDocument service implements the DrawingDocument service. This means that every presentation document looks like a drawing document. To distinguish between the two document types, you must first check for a presentation (Impress) document and then check for a drawing document. OOME 4.1. Pg 562

## Draw

Things pertaining to **Draw**.

- For future reference:  The PresentationDocument service implements the DrawingDocument service. This means that every presentation document looks like a drawing document. To distinguish between the two document types, you must first check for a presentation (Impress) document and then check for a drawing document. OOME 4.1. Pg 562
