PortrayalCatalogueValidator
===========================
Specification checks implemented:

1. All StyleIDs and ColorIDs begin with their respective letters.

2. All StyleIDs and ColorIDs are unique.
-> This check is only for IDs which begins with correct letters

3. Sheets with Color column(s) have valid colors (corresponds with the Colors Sheet)
-> Only RasterSymbolizer sheet have no color column
-> TextSymbolizer sheet have only 1 color column while the rest have 2

4. Implement format check for Colors Sheet sRGB column
-> Values must begin with '#'
-> Values must be Hexadecimal (from position 1)

5. Each line should contain values (no cell grouping).
-> Once a merged cell is found, message will be displayed and the rest of the checks will not run
-> Currently only implemented on PointSymbolizer Sheet. Not sure to duplicate code into other sheets or make function take in sheets as variables

6. Cells shouldn’t contain newlines (no line breaks).
-> Two types of line breaks
	a. Null row
	b. Empty row
-> Once a line break is found, message will be displayed and the rest of the checks will not run
-> Currently only implemented on PointSymbolizer Sheet. Not sure to duplicate code into other sheets or make function take in sheets as variables


Other notes:

1. Only data rows from Colors sheet are being read in and store as objects

2. Applications can only accept.xlsx extension. But could not handle wrong .xlsx file.
