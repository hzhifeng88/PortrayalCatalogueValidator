import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedList;
import java.util.StringTokenizer;

import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ValidatePointSymbolizer {

	private Sheet pointSheet;
	private ColorsObject cObject;
	private LinkedList<ColorsObject> storeColorsObjectList = new LinkedList<ColorsObject>();
	private ArrayList<String> storeMergedCells = new ArrayList<String>();
	private ArrayList<String> storeEmptyRows = new ArrayList<String>();
	private ArrayList<String> storeRightStyleID = new ArrayList<String>();
	private ArrayList<String> storeWrongStyleID = new ArrayList<String>();
	private ArrayList<String> storeDuplicateStyleID = new ArrayList<String>();
	private ArrayList<String> storeInvalidColorCells = new ArrayList<String>();
	private ArrayList<String> storeMissingValueCells = new ArrayList<String>();
	private HTMLEditorKit kit;
	private HTMLDocument doc;

	public ValidatePointSymbolizer(Sheet pointSheet,
			LinkedList<ColorsObject> list, HTMLEditorKit kit, HTMLDocument doc) {

		this.pointSheet = pointSheet;
		this.kit = kit;
		this.doc = doc;
		storeColorsObjectList = list;

		checkMergedCells();
		checkEmptyRows();

		if (storeMergedCells.isEmpty() == true
				&& storeEmptyRows.isEmpty() == true) {

			checkStyleID();
			checkColorValid();
			checkMissingAttributes();

			if (storeWrongStyleID.isEmpty() == false
					|| storeInvalidColorCells.isEmpty() == false
					|| storeDuplicateStyleID.isEmpty() == false) {
				printAllError();
			}

		} else {
			
			try {
				
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Error Sheet: <font color=#ED0E3F>PointSymbolizer</font color></font>", 0, 0, null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#088542>-------------------------------------- </font color></font>", 0, 0, null);

				if (storeMergedCells.isEmpty() == false) {
				
					kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>Merged cells found! Please correct this and try again.</font color></font>", 0, 0,null);
					kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeMergedCells + "</font color></font>", 0, 0, null);

				}
				
				if (storeEmptyRows.isEmpty() == false) {
					
					kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Empty rows found! Please correct this and try again.</font color></font>", 0, 0,null);
					kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Row number: <font color=#ED0E3F>" + storeEmptyRows + "</font color></font>", 0, 0, null);

				}
				
			} catch (BadLocationException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public ArrayList<String> checkMergedCells() {

		String cellNumber;

		for (int count = 0; count < pointSheet.getNumMergedRegions(); count++) {

			cellNumber = "";

			String tempString = pointSheet.getMergedRegion(count).toString()
					.substring(41);
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
		Collections.sort(storeMergedCells);
		return storeMergedCells;
	}

	public void checkEmptyRows() {

		boolean isRowEmpty = false;

		for (int rowCount = 4; rowCount <= pointSheet.getLastRowNum(); rowCount++) {

			isRowEmpty = false;
			Row row = pointSheet.getRow(rowCount);

			if (row == null) {

				System.out.println("Null row: "
						+ Integer.toString(rowCount + 1));

				storeEmptyRows.add(Integer.toString(rowCount + 1));
				continue;

			}
			
			// Check if all cells are empty
			for (int cellCount = 0; cellCount < row.getLastCellNum(); cellCount++) {

				if (row.getCell(cellCount) == null
						|| row.getCell(cellCount).toString().trim().equals("")) {

					isRowEmpty = true;

				} else {

					isRowEmpty = false;
					break;
				}
			}
			if (isRowEmpty == true) {

				storeEmptyRows.add(Integer.toString(rowCount + 1));
				System.out.println("Blank row: "
						+ Integer.toString(rowCount + 1));

			}
		}
	}

	// This function directly reads in the rows from the sheet and performs the
	// specification checks. No data structures are used to store the data
	public void checkStyleID() {

		boolean wrongStyleID = false;
		boolean foundDuplicate = false;

		// if two rows cells are merged, i.e row 3 & 4, row 3 will be the one
		// with text and row 4 will be ""
		for (int rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {

			wrongStyleID = false;
			foundDuplicate = false;

			Row row = pointSheet.getRow(rowIndex);

			String tempString = row.getCell(5).toString();

			if (!tempString.equalsIgnoreCase("")) {

				char firstChar = tempString.charAt(0);

				// Ensure begin with 'P')
				if (firstChar != 'P') {
					wrongStyleID = true;
					storeWrongStyleID.add("F" + Integer.toString(rowIndex + 1));
				}

				if (wrongStyleID == false) {

					if (storeRightStyleID.isEmpty() == true) {
						storeRightStyleID.add(tempString);
					} else {

						for (int count = 0; count < storeRightStyleID.size(); count++) {

							if (storeRightStyleID.get(count).equalsIgnoreCase(
									tempString)) {

								storeDuplicateStyleID.add("F"
										+ Integer.toString(rowIndex + 1));
								foundDuplicate = true;
								break;

							}
						}
						if (foundDuplicate == false) {
							storeRightStyleID.add(tempString);
						}
					}
				}
			}
		}
	}

	// Column P and X store color values or references
	public void checkColorValid() {

		int rowIndex;

		for (rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {

			Row row = pointSheet.getRow(rowIndex);

			matchColor(row.getCell(15).toString(), 15, rowIndex);
			matchColor(row.getCell(23).toString(), 23, rowIndex);

		}
		Collections.sort(storeInvalidColorCells);
	}

	public void matchColor(String tempStringColor, int columnNum, int rowIndex) {

		// Check if each color is valid in Colors Sheet
		for (int count = 0; count < storeColorsObjectList.size(); count++) {

			cObject = storeColorsObjectList.get(count);

			if (tempStringColor.equalsIgnoreCase(cObject.getColorID())
					|| tempStringColor.equalsIgnoreCase(cObject.getsRGB())) {

				return;
			}
		}

		if (columnNum == 15) {
			storeInvalidColorCells.add("P" + Integer.toString(rowIndex + 1));
		} else if (columnNum == 23) {
			storeInvalidColorCells.add("X" + Integer.toString(rowIndex + 1));
		}
	}

	public void checkMissingAttributes(){
		
		for (int rowCount = 4; rowCount <= pointSheet.getLastRowNum(); rowCount++) {
			
			for(int emptyRowCount = 0; emptyRowCount < storeEmptyRows.size(); emptyRowCount++){
				
				if(rowCount == Integer.parseInt(storeEmptyRows.get(emptyRowCount).toString())){
					break;
				}
			}
			
			Row row = pointSheet.getRow(rowCount);

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
		Collections.sort(storeMissingValueCells);
	}
	
	public void printAllError() {

		try {
			
			kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Error Sheet: <font color=#ED0E3F><b>PointSymbolizer</b></font color></font>", 0, 0, null);
			kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#088542>-------------------------------------- </font color></font>", 0, 0, null);

			if (storeWrongStyleID.isEmpty() == false) {
			
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>Style ID does not begin with 'P'</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeWrongStyleID + "</font color></font>", 0, 0, null);

			}
			
			if(storeDuplicateStyleID.isEmpty() == false){
				
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Duplicate Style ID</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeDuplicateStyleID + "</font color></font>", 0, 0, null);

			}
			
			if (storeInvalidColorCells.isEmpty() == false) {
				
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Color is invalid (Rule 1: Begins with '#', Rule 2: 6 hexadecimal representation)</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeInvalidColorCells + "</font color></font>", 0, 0, null);

			}
			
			if (storeMissingValueCells.isEmpty() == false) {
				
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Missing values found (Mandatory)</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeMissingValueCells + "</font color></font>", 0, 0, null);

			}
		} catch (BadLocationException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
