import java.util.ArrayList;

import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ValidatePointSymbolizer extends CommonValidator{

	private Sheet pointSheet;

	public ValidatePointSymbolizer(Sheet pointSheet, Workbook originalWorkbook, ArrayList<String> list, HTMLEditorKit kit, HTMLDocument doc) {
		
		super(pointSheet,originalWorkbook, list, kit, doc);
		this.pointSheet = pointSheet;
		
		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
			checkModifiedHeader();
			performChecks();
			printValueError();
		} 
	}

	public void performChecks() {

		for (int rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {

			Row row = pointSheet.getRow(rowIndex);
			
			// Check valid ID and duplicate
			checkIDAndDuplicate('P', "F", rowIndex, 5);
			
			// Check color valid
			matchColor(row.getCell(15).toString(), "P", rowIndex);
			matchColor(row.getCell(23).toString(), "X", rowIndex);
			
			// Check missing attributes
			checkMissingAttributes(row, rowIndex);
		}
	}
}
