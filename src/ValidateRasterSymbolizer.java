import java.io.IOException;
import java.util.ArrayList;

import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ValidateRasterSymbolizer {

	private Sheet rasterSheet;
	private ArrayList<String> storeRightStyleID = new ArrayList<String>();						//StyleIDs which begins with right letters
	private ArrayList<String> storeWrongStyleID = new ArrayList<String>();						//StyleIDs which begins with wrong letters
	private ArrayList<String> storeDuplicateStyleID = new ArrayList<String>();					//StyleIDs which are duplicates
	private HTMLEditorKit kit;
	private HTMLDocument doc;

	public ValidateRasterSymbolizer(Sheet sheet, HTMLEditorKit kit, HTMLDocument doc) {

		this.rasterSheet = sheet;
		this.kit = kit;
		this.doc = doc;
		
		checkStyleID();
		
		if (storeWrongStyleID.isEmpty() == false
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
		for (int rowIndex = 4; rowIndex <= rasterSheet.getLastRowNum(); rowIndex++) {

			wrongStyleID = false;
			foundDuplicate = false;
			
			Row row = rasterSheet.getRow(rowIndex);

			String tempString = row.getCell(5).toString();

			if (!tempString.equalsIgnoreCase("")) {

				char firstChar = tempString.charAt(0);

				// Ensure begins with 'R')
				if (firstChar != 'R') {
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
	
	public void printAllError() {

	try {
			
			kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4><br>Error Sheet: <font color=#ED0E3F><b>RasterSymbolizer</b></font color></font>", 0, 0, null);
			kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#088542>-------------------------------------- </font color></font>", 0, 0, null);

			if (storeWrongStyleID.isEmpty() == false) {
			
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>Style ID does not begin with 'R'</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeWrongStyleID + "</font color></font>", 0, 0, null);

			}
			
			if(storeDuplicateStyleID.isEmpty() == false){
				
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Duplicate Style ID</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeDuplicateStyleID + "</font color></font>", 0, 0, null);

			}
			
			
		} catch (BadLocationException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}
