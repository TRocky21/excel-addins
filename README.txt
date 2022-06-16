Excel Add-Ins
Written by Torin Rockwell
Version 20220614
https://github.com/TRocky21/excel-addins


AddTemplateSheet:
	Opens a dialog box in the default templates folder and places selected sheet template after the active sheet.

AutofitColumns:
	If one cell is selected, autofit all columns in the active worksheet.
	If multiple cells are selected, autofit all columns in the selected range.

BackupWorkbook:
	Saves a timestamped copy of the workbook to "Documents/Excel Backups" folder.
	Creates folder if it does not exist.

BatchLink:
	If a cell value in the current selection matches a sheet name, link the cell to A1 in the corresponding sheet.

CellLink:
	Opens a userform that allows user to create a button or cell link to anywhere within the workbook.

CenterAcross:
	Applies "Center Across Selection" formatting to selected cells.

ContentsLinks:
	Starting in the top cell of the selection, creates links to cell A1 of every sheet in the workbook.
	Does not work if cells in multiple columns are selected.

CopyPaste:
	If one region is selected, copies the row of cells immediately above and pastes them into selection.
	If multiple regions are selected, copies the first selected region and pastes it into subsequent regions.

CreateTable:
	If current selection does not intersect with another table, creates a table.
	If an entire table is selected, removes formatting and unlists the table.

CycleCase:
	Cycles text case between lower, proper, and upper based on the first cell in the selected region.

DefineWord:
	Gets the definition of the word in the selected cell.
	Worksheet function "=DEFINE(word as String)" will return the definition of a word.

FindReplace:
	Replacement for built-in Find and Replace function.
	Adds ability to search within selection, sheet, or workbook.
	Adds ability to search in formulas only, values only, or formulas and values.

FlipRange:
	If selection contains only one row or column, flips the values in that row or column.

FormulasToValues:
	Convert selected cell formulas to values.

GoHome:
	If the top left cell is selected, select the first sheet in the workbook.
	Else, select the top left cell in the current sheet.

IfError:
	If cells contain formulas and do not have an IFERROR statement, wraps them in IFERROR with specified text.
	Can look within selection, sheet, or workbook.

PasswordGenerator:
	Generate a password and copy it to the clipboard.

PasteImage:
	Shortcut to paste from clipboard as an image.

ResetCellSize:
	Reset selected cell(s) to default size.

ResetSelections:
	In every sheet in the workbook, select cell A1 and scroll view to top left.

SelectBlanks:
	Select all blank cells within selection.

SwapValues:
	Swaps the values of any two selected cells.
	If two ranges of equal size are selected, swaps their values.

TransposeRange:
	Switches the rows and columns of the selection.

TravelInformation:
	Creates two Excel formula functions, TRAVELTIME and TRAVELDISTANCE
	TRAVELTIME takes two arguments, origin and destination, and returns the travel time in seconds.
	TRAVELDISTANCE takes two arguments, origin and destination, and returns the travel distance in meters.

UnhideSheets:
	If any sheets in the workbook are hidden, unhide them.

UnitConverter:
	Convert between units.