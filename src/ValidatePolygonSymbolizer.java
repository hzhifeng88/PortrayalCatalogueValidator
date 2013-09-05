import java.util.ArrayList;

import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.*;

public class ValidatePolygonSymbolizer extends CommonValidator{

	private Sheet polygonSheet;
	private boolean sheetCorrect = false;

	public ValidatePolygonSymbolizer(Sheet sheet, Workbook originalWorkbook, ArrayList<String> list , HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, list, kit, doc);
		this.polygonSheet = sheet;
		
		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
			checkModifiedHeader();
			performChecks();
			printValueError();
			
			if(getHasError() == false){
				sheetCorrect = true;
			}else {
				sheetCorrect = false;
			}
		}
	}
	
	public boolean isSheetCorrect() {
		return sheetCorrect;
	}
	
	public void performChecks() {

		for (int rowIndex = 4; rowIndex <= polygonSheet.getLastRowNum(); rowIndex++) {

			Row row = polygonSheet.getRow(rowIndex);
			
			// Check valid ID and duplicate
			checkIDAndDuplicate('A', "F", rowIndex, 5);
			
			// Check color valid
			if(row.getCell(12) != null && !row.getCell(12).toString().equalsIgnoreCase("")) {
				matchColor(row.getCell(12).toString(), "M", rowIndex);
			}
			if(row.getCell(18) != null && !row.getCell(18).toString().equalsIgnoreCase("")) {
				matchColor(row.getCell(18).toString(), "S", rowIndex);
			}
			
			// Check missing attributes
			checkMissingAttributes(row, rowIndex);
		}
	}
}
