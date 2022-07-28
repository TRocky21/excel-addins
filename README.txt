Excel Add-Ins
Written by Torin Rockwell
Version 20220713
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

ChangeCase:
	Allows user to change case of selection to lower, sentence, proper, or upper.

ContentsLinks:
	Starting in the top cell of the selection, creates links to cell A1 of every sheet in the workbook.
	Does not work if cells in multiple columns are selected.

CopyPaste:
	If one region is selected, copies the row of cells immediately above and pastes them into selection.
	If multiple regions are selected, copies the first selected region and pastes it into subsequent regions.

CountErrors:
	Loop through each sheet and display count of errors in active workbook.

CreateTable:
	If current selection does not intersect with another table, creates a table.
	If an entire table is selected, removes formatting and unlists the table.

CycleCase:
	Cycles text case between lower, proper, and upper based on the first cell in the selected region.

DatePicker:
	Opens a date picker window to allow the user to put a chosen date in the selected cell.

DecodeFastenerNumber:
	Decode a fastener part number to determine its characteristics and specifications.

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

GenerateFastenerNumber:
	Generate a fastener part number based on its characteristics and specifications.

GoHome:
	If the top left cell is selected, select the first sheet in the workbook.
	Else, select the top left cell in the current sheet.

IfError:
	If cells contain formulas and do not have an IFERROR statement, wraps them in IFERROR with specified text.
	Can look within selection, sheet, or workbook.

MoveShapeToRange:
	Allows user to move a shape to a selected range and set its dimensions equal to those of the range.

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

SelectionSummary:
	Provides a summary of the cells in the selected range.
	Count, unique values, sum, average, etc.
	Second tab provides statistical analysis of range, e.g. median, mode, skew, kurtosis, confidence, etc.

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

UserDefinedFunctions:
	Compiles miscellaneous UDFs.
	ISEMAIL - validate email addresses.
	TEXTSPLIT - split text string by a delimiter.
	COUNTUNIQUE - return number of unique values in a range.
	REVERSE - reverses a given string.
	DATE_ADD - access VBA DateAdd function in Excel.
	ROUNDSIGFIGS - round a given number to specified significant figures.
	ISPRIME - determine if a given number is prime.