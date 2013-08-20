import java.util.*;
import java.io.IOException;

import javax.swing.text.BadLocationException;
import javax.swing.text.html.*;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class CommonValidator {

	private char idAlphabet;
	public boolean hasError = false;
	private static String hexadecimal = "0123456789abcdefABCDEF";
	private Sheet sheet;
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

	
	public CommonValidator(Sheet sheet, HTMLEditorKit kit, HTMLDocument doc){
		
		this.sheet = sheet;	
		this.kit = kit;
		this.doc = doc;
	}
	
	public CommonValidator(Sheet sheet, ArrayList<String> list, HTMLEditorKit kit, HTMLDocument doc){
		
		this.sheet = sheet;	
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
	
	public void checkIDAndDuplicate(char idAlphabet, String column, int rowIndex, int columnIndex){
		
		boolean wrongStyleID = false;
		boolean foundDuplicate = false;
		this.idAlphabet = idAlphabet;

		Row row = sheet.getRow(rowIndex);

		String tempString = row.getCell(columnIndex).toString();

		if (!tempString.equalsIgnoreCase("")) {

			char firstChar = tempString.charAt(0);

			if (firstChar != idAlphabet) {
				wrongStyleID = true;
				storeWrongID.add(column + Integer.toString(rowIndex + 1));
			}

			if (wrongStyleID == false) {

				if (storeRightID.isEmpty() == true) {
					storeRightID.add(tempString);
				} else {

					for (int count = 0; count < storeRightID.size(); count++) {

						if (storeRightID.get(count).equalsIgnoreCase(tempString)) {
							storeDuplicateID.add(column + Integer.toString(rowIndex + 1));
							foundDuplicate = true;
						}
					}
					
					if (foundDuplicate == false) {
						storeRightID.add(tempString);
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

		for (int count = 0; count < storeColorID.size(); count++) {
			
			if(tempStringColor.equalsIgnoreCase(storeColorID.get(count))){
				return;
			}else if(isRGB(tempStringColor) == true){
				return;
			}
		}
		storeInvalidColorCells.add(currentColumn + Integer.toString(rowIndex + 1));
	}
	
	public void checkMissingAttributes(){
		
		for (int rowCount = 4; rowCount <= sheet.getLastRowNum(); rowCount++) {
			
			for(int emptyRowCount = 0; emptyRowCount < storeEmptyRows.size(); emptyRowCount++){
				
				if(rowCount == Integer.parseInt(storeEmptyRows.get(emptyRowCount).toString())){
					break;
				}
			}
			
			Row row = sheet.getRow(rowCount);

			// Checks for mandatory columns here
			if(row.getCell(0) == null){
				storeMissingValueCells.add("A" + Integer.toString(rowCount + 1));
			}
			
			if(row.getCell(2) == null){
				storeMissingValueCells.add("C" + Integer.toString(rowCount + 1));
			}
			
			if(row.getCell(3) == null){
				storeMissingValueCells.add("D" + Integer.toString(rowCount + 1));
			}
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