import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ValidatePointSymbolizer extends CommonValidator{

	private Sheet pointSheet;

	public ValidatePointSymbolizer(Sheet pointSheet, ArrayList<String> list, HTMLEditorKit kit, HTMLDocument doc) {
		
		super(pointSheet,list, kit, doc);
		this.pointSheet = pointSheet;
		
		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
			checkStyleID();
			checkColorValid();
			checkMissingAttributes();
			printValueError();
		} 
	}

	public void checkStyleID() {

		for (int rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {

			checkIDAndDuplicate('P', "F", rowIndex, 5);
		}
	}

	// Column P and X store color values or references
	public void checkColorValid() {

		int rowIndex;
		
		for (rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {

			Row row = pointSheet.getRow(rowIndex);
			matchColor(row.getCell(15).toString(), "P", rowIndex);
			matchColor(row.getCell(23).toString(), "X", rowIndex);
		}
	}
}
