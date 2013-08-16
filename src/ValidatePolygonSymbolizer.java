import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedList;

import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ValidatePolygonSymbolizer {

	private Sheet polygonSheet;
	private ColorsObject cObject;
	private LinkedList<ColorsObject> storeColorsObjectList = new LinkedList<ColorsObject>();
	private ArrayList<String> storeRightStyleID = new ArrayList<String>();						//StyleIDs which begins with right letters
	private ArrayList<String> storeWrongStyleID = new ArrayList<String>();						//StyleIDs which begins with wrong letters
	private ArrayList<String> storeDuplicateStyleID = new ArrayList<String>();					//StyleIDs which are duplicates
	private ArrayList<String> storeInvalidColorCells = new ArrayList<String>();					//Color values/codes are invalid(not found)
	private HTMLEditorKit kit;
	private HTMLDocument doc;

	public ValidatePolygonSymbolizer(Sheet sheet, LinkedList<ColorsObject> list , HTMLEditorKit kit, HTMLDocument doc) {

		this.polygonSheet = sheet;
		this.kit = kit;
		this.doc = doc;
		storeColorsObjectList = list;

		checkStyleID();
		checkColorValid();

		if (storeWrongStyleID.isEmpty() == false
				|| storeInvalidColorCells.isEmpty() == false
				|| storeDuplicateStyleID.isEmpty() == false) {
			printAllError();
		}
	}

	// This function directly reads in the rows from the sheet and performs the
	// specification checks. No data structures are used to store the data
	public void checkStyleID() {

		boolean wrongStyleID = false;
		boolean foundDuplicate = false;
		
		// if two rows cells are merged, i.e row 3 & 4, row 3 will be the one
		// with text and row 4 will be ""
		for (int rowIndex = 4; rowIndex <= polygonSheet.getLastRowNum(); rowIndex++) {

			wrongStyleID = false;
			foundDuplicate = false;
			
			Row row = polygonSheet.getRow(rowIndex);

			String tempString = row.getCell(5).toString();

			if (!tempString.equalsIgnoreCase("")) {

				char firstChar = tempString.charAt(0);

				// Ensure begins with 'A')
				if (firstChar != 'A') {
					wrongStyleID = true;
					storeWrongStyleID.add("F"
							+ Integer.toString(rowIndex + 1));
				}
				
				if(wrongStyleID == false){
					
					if(storeRightStyleID.isEmpty() == true){
						storeRightStyleID.add(tempString);
					}else{
						
						for(int count = 0; count < storeRightStyleID.size(); count++){
							
							if(storeRightStyleID.get(count).equalsIgnoreCase(tempString)){

								storeDuplicateStyleID.add("F" + Integer.toString(rowIndex + 1));
								foundDuplicate = true;
								break;
								
							} 
						}
						if(foundDuplicate == false){
							storeRightStyleID.add(tempString);
						}
					}
				}
			}
		}
	}
	
	// Column K and Q store color values or references
		public void checkColorValid() {

			int rowIndex;

			for (rowIndex = 4; rowIndex <= polygonSheet.getLastRowNum(); rowIndex++) {

				Row row = polygonSheet.getRow(rowIndex);

				matchColor(row.getCell(10).toString(), 10, rowIndex);
				matchColor(row.getCell(16).toString(), 16, rowIndex);

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

			if (columnNum == 10) {
				storeInvalidColorCells.add("K" + Integer.toString(rowIndex + 1));
			} else if (columnNum == 16) {
				storeInvalidColorCells.add("Q" + Integer.toString(rowIndex + 1));
			}
		}

		public void printAllError() {

			try {
				
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4><br>Error Sheet: <font color=#ED0E3F><b>PolygonSymbolizer</b></font color></font>", 0, 0, null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#088542>-------------------------------------- </font color></font>", 0, 0, null);

				if (storeWrongStyleID.isEmpty() == false) {
				
					kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>Style ID does not begin with 'A'</font color></font>", 0, 0,null);
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
			} catch (BadLocationException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
}
