import java.util.*;
import java.io.IOException;

import javax.swing.text.BadLocationException;
import javax.swing.text.html.*;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class CommonValidator {

	private char idAlphabet;
	public boolean hasError = false;
	private static String hexadecimal = "0123456789abcdefABCDEF";
	private Sheet sheet;
	private Workbook originalWorkbook;
	private HTMLEditorKit kit;
	private HTMLDocument doc;
	private ArrayList<String> storeRightID = new ArrayList<String>();
	private ArrayList<String> storeWrongID = new ArrayList<String>();
	private ArrayList<String> storeDuplicateID = new ArrayList<String>();
	private ArrayList<String> storeColorID = new ArrayList<String>();
	private ArrayList<String> storeMergedCells = new ArrayList<String>();
	private ArrayList<String> storeEmptyRows = new ArrayList<String>();
	private ArrayList<String> storeMissingValueCells = new ArrayList<String>();
	private ArrayList<String> storeInvalidColorCells = new ArrayList<String>();
	private ArrayList<String> storeLineBreakCells = new ArrayList<String>();
	private ArrayList<String> storeModifiedHeaderCells = new ArrayList<String>();
	
	public CommonValidator(Sheet sheet, Workbook originalWorkbook, HTMLEditorKit kit, HTMLDocument doc){
		
		this.sheet = sheet;	
		this.originalWorkbook = originalWorkbook;
		this.kit = kit;
		this.doc = doc;
	}
	
	public CommonValidator(Sheet sheet, Workbook originalWorkbook, ArrayList<String> list, HTMLEditorKit kit, HTMLDocument doc){
		
		this.sheet = sheet;
		this.originalWorkbook = originalWorkbook;
		this.kit = kit;
		this.doc = doc;
		storeColorID = list;
	}
	
	public void checkMergedCells() {

		String cellNumber;

		for (int count = 0; count < sheet.getNumMergedRegions(); count++) {

			cellNumber = "";

			String tempString = sheet.getMergedRegion(count).toString().substring(41);
			StringTokenizer tokenizer = new StringTokenizer(tempString, ":");

			String cell = tokenizer.nextToken();

			for (int count1 = 0; count1 < cell.length(); count1++) {

				char checkChar = cell.charAt(count1);

				if (Character.isDigit(checkChar)) {
					cellNumber = cellNumber.concat(String.valueOf(checkChar));
				}

			}
			// Begin check from row 5 onwards
			if (Integer.parseInt(cellNumber) > 4) {
				storeMergedCells.add(cell);
			}
		}
	}
	
	public void checkEmptyRows() {

		boolean isRowEmpty = false;

		for (int rowCount = 4; rowCount <= sheet.getLastRowNum(); rowCount++) {

			isRowEmpty = false;
			Row row = sheet.getRow(rowCount);

			if (row == null) {
				System.out.println("Null row: " + Integer.toString(rowCount + 1));
				storeEmptyRows.add(Integer.toString(rowCount + 1));
				continue;
			}
			
			// Check if all cells are empty
			for (int cellCount = 0; cellCount < row.getLastCellNum(); cellCount++) {

				if (row.getCell(cellCount) == null || row.getCell(cellCount).toString().trim().equals("")) {
					isRowEmpty = true;
				} else {
					isRowEmpty = false;
					break;
				}
			}
			if (isRowEmpty == true) {
				storeEmptyRows.add(Integer.toString(rowCount + 1));
				System.out.println("Blank row: " + Integer.toString(rowCount + 1));
			}
		}
	}
	
	public String columnIndexToLetter(int columnIndex) { 
  
		int base = 26;   
		StringBuffer b = new StringBuffer(); 
		
		do {  
			int digit = columnIndex % base + 65;  
			b.append(Character.valueOf((char) digit));  
			columnIndex = (columnIndex / base) - 1; 
			
		} while (columnIndex >= 0);   
		
		return b.reverse().toString();
	}
	
	public void checkModifiedHeader(){
		
		Sheet originalSheet = originalWorkbook.getSheet(sheet.getSheetName());
		
		for(int rowIndex = 0; rowIndex < 4; rowIndex++){
			
			Row row = sheet.getRow(rowIndex);
			Row originalRow = originalSheet.getRow(rowIndex);
			
			for(int columnIndex = 0; columnIndex < row.getLastCellNum(); columnIndex++){
				
				if(row.getCell(columnIndex) == null && originalRow.getCell(columnIndex) == null){
					continue;
				}

				if(row.getCell(columnIndex) != null && originalRow.getCell(columnIndex) == null){
					storeModifiedHeaderCells.add(columnIndexToLetter(columnIndex) + Integer.toString(rowIndex + 1));
					continue;
				}
						
				if(row.getCell(columnIndex).toString().equalsIgnoreCase(originalRow.getCell(columnIndex).toString())){
					continue;
				}else {
					storeModifiedHeaderCells.add(columnIndexToLetter(columnIndex) + Integer.toString(rowIndex + 1));
				}
			}
		}
	}

	public void checkLineBreak(String tempString, String column, int rowIndex){

		if(tempString.contains("\n")){
			storeLineBreakCells.add(column + Integer.toString(rowIndex + 1));
		}
	}
	
	public void checkIDAndDuplicate(char idAlphabet, String column, int rowIndex, int columnIndex){
		
		boolean wrongStyleID = false;
		this.idAlphabet = idAlphabet;

		Row row = sheet.getRow(rowIndex);

		if(row.getCell(columnIndex) != null){
			
			String tempString = row.getCell(columnIndex).toString();

			if (!tempString.equalsIgnoreCase("")) {
				
				checkLineBreak(tempString, column, rowIndex);

				char firstChar = tempString.charAt(0);

				if (firstChar != idAlphabet) {
					wrongStyleID = true;
					storeWrongID.add(column + Integer.toString(rowIndex + 1));
				}

				if (wrongStyleID == false) {

					if (storeRightID.isEmpty() == true) {
						storeRightID.add(tempString);
					} else {

						if(storeRightID.contains(tempString)){
							storeDuplicateID.add(column + Integer.toString(rowIndex + 1));
						}else{
							storeRightID.add(tempString);
						}
					}
				}
			}
		}
	}
	
	public boolean isRGB(String tempStringColor) {	
		
		if (tempStringColor.charAt(0) != '#'){
			return false;	
		}
		if (tempStringColor.length() != 7){
			return false;	
		}
		
		for (int stringIndex = 1; stringIndex < tempStringColor.length(); stringIndex ++) {
			
			if (hexadecimal.indexOf(tempStringColor.charAt(stringIndex)) == -1) 
				return false; 		
		}		
		return true;		
	}
	
	public void matchColor(String tempStringColor, String currentColumn, int rowIndex) {

		checkLineBreak(tempStringColor, currentColumn, rowIndex);
		
		if(storeColorID.contains(tempStringColor)){
			return;
		}else if(isRGB(tempStringColor) == true){
			return;
		}
		storeInvalidColorCells.add(currentColumn + Integer.toString(rowIndex + 1));
	}
	
	public void checkMissingAttributes(Row row, int rowIndex){
		
		// Checks for mandatory columns here
		if(row.getCell(0) == null || row.getCell(2).toString().equalsIgnoreCase("")){
			storeMissingValueCells.add("A" + Integer.toString(rowIndex + 1));
		}
			
		if(row.getCell(2) == null || row.getCell(2).toString().equalsIgnoreCase("")){
			storeMissingValueCells.add("C" + Integer.toString(rowIndex + 1));
		}
			
		if(row.getCell(3) == null || row.getCell(2).toString().equalsIgnoreCase("")){
			storeMissingValueCells.add("D" + Integer.toString(rowIndex + 1));
		}
	}

	public boolean printFormatError(){

		try {
			
			if(sheet.getSheetName().equalsIgnoreCase("PointSymbolizer")){
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Error Sheet: <font color=#ED0E3F><b>" + sheet.getSheetName() + "</b></font color></font>", 0, 0, null);
			}else{
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4><br>Error Sheet: <font color=#ED0E3F><b>" + sheet.getSheetName() + "</b></font color></font>", 0, 0, null);
			}
			kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#088542>-------------------------------------- </font color></font>", 0, 0, null);
		
			if (storeMergedCells.isEmpty() == true && storeEmptyRows.isEmpty() == true) {
				return false;
			}else {

				if (storeMergedCells.isEmpty() == false) {
					hasError = true;
					Collections.sort(storeMergedCells);
					kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>Merged cells found! Please correct this and try again.</font color></font>", 0, 0,null);
					kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeMergedCells + "</font color></font>", 0, 0, null);
				}
					
				if (storeEmptyRows.isEmpty() == false) {
					hasError = true;
					kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Empty rows found! Please correct this and try again.</font color></font>", 0, 0,null);
					kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Row number: <font color=#ED0E3F>" + storeEmptyRows + "</font color></font>", 0, 0, null);
				}		
			} 
		} catch (BadLocationException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return true;
	}
	
	public void printValueError(){
		
		try {
			
			if (storeWrongID.isEmpty() == false) {
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>ID does not begin with '" + idAlphabet + "'</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeWrongID + "</font color></font>", 0, 0, null);
			}
			
			if(storeDuplicateID.isEmpty() == false){
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Duplicate ID</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeDuplicateID + "</font color></font>", 0, 0, null);
			}

			if (storeInvalidColorCells.isEmpty() == false) {
				hasError = true;
				Collections.sort(storeInvalidColorCells);
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Color is invalid (Rule 1: Begins with '#', Rule 2: 6 hexadecimal representation)</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeInvalidColorCells + "</font color></font>", 0, 0, null);
			}	
			
			if (storeMissingValueCells.isEmpty() == false) {
				hasError = true;
				Collections.sort(storeMissingValueCells);
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Missing values found (Mandatory)</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeMissingValueCells + "</font color></font>", 0, 0, null);
			}
			
			if(storeLineBreakCells.isEmpty() == false){
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Cell contains line break</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeLineBreakCells + "</font color></font>", 0, 0, null);
			}
			
			if(storeModifiedHeaderCells.isEmpty() == false){
				hasError = true;
				Collections.sort(storeModifiedHeaderCells);
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Header cells are modified! </font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeModifiedHeaderCells + "</font color></font>", 0, 0, null);
			}
			
			if(hasError == false){
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> No error found!! </font color></font>", 0, 0,null);
			}
		} catch (BadLocationException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
