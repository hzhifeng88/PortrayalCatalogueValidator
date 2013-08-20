PortrayalCatalogueValidator
===========================
Specifications implemented: 

1. All Style IDs and Color IDs begin with their respective letters and are unique. 

2. Sheets with Color column(s) have valid colors (corresponds with the Colors Sheet). 

3. Implement format check for Colors Sheet sRGB column. 

* Values must begin with '#' 
* Have a length of 7 (inclusive of '#') 
* Values must be Hexadecimal (from position 1) 

4. Each line should contain values (no cell grouping). 

5. Cells shouldn’t contain newlines (no line breaks). 

6. Checks for mandatory column. 

7. Sheets should not contains any empty rolls (null or blank). 


Other notes:

1. Applications can only accept.xlsx extension. But could not handle wrong .xlsx file.
